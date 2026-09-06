using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Tests;

/// <summary>
/// Builds small in-memory documents and inspects what a round-trip produced.
/// </summary>
/// <remarks>
/// Tests assert on the saved package rather than on the model, because the failures these cover were
/// all cases where the model looked right and the saved document did not.
/// </remarks>
internal static class DocumentFixture
{
    public const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>A 1x1 transparent PNG, small enough to inline.</summary>
    public static byte[] TinyPng { get; } = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/Z1kAAAAASUVORK5CYII=");

    /// <summary>Builds a paragraph, optionally styled as a heading.</summary>
    public static string Paragraph(string text, int headingLevel = 0) =>
        $"<w:p>{(headingLevel > 0 ? $"<w:pPr><w:pStyle w:val='Heading{headingLevel}'/></w:pPr>" : "")}" +
        $"<w:r><w:t>{text}</w:t></w:r></w:p>";

    /// <summary>Builds a single-cell table wrapping the given block content.</summary>
    public static string Table(string content) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w='2000'/></w:tblGrid>" +
        $"<w:tr><w:tc>{content}</w:tc></w:tr></w:tbl>";

    /// <summary>Creates a .docx package with the given body content.</summary>
    public static byte[] Create(string bodyContent, Action<WordprocessingDocument>? configure = null)
    {
        using var buffer = new MemoryStream();

        using (var package = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            package.AddMainDocumentPart().Document =
                new Document(new Body($"<w:body xmlns:w='{W}' xmlns:r='{R}'>{bodyContent}</w:body>"));
            configure?.Invoke(package);
        }

        return buffer.ToArray();
    }

    /// <summary>Parses a package into a document tree.</summary>
    public static WordDocument Parse(byte[] package)
    {
        using var stream = new MemoryStream(package);
        using var parser = new WordDocumentTreeParser();
        return parser.ParseFromStream(stream);
    }

    /// <summary>Saves a document and returns the resulting package bytes.</summary>
    public static byte[] Save(WordDocument document) => document.ToDocxBytes();

    /// <summary>Saves a document and returns the body XML of the result.</summary>
    public static string SavedBodyXml(WordDocument document)
    {
        using var stream = new MemoryStream(Save(document));
        using var package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart!.Document.Body!.OuterXml;
    }

    /// <summary>Saves a document and returns its concatenated visible text.</summary>
    public static string SavedText(WordDocument document) => TextOf(SavedBodyXml(document));

    /// <summary>Concatenates every <c>w:t</c> in a fragment of body XML.</summary>
    public static string TextOf(string xml) =>
        string.Concat(XDocument.Parse(xml).Descendants(XName.Get("t", W)).Select(e => e.Value));

    /// <summary>Opens a saved package for inspection. Dispose the returned holder.</summary>
    public static SavedPackage Open(byte[] package) => new(package);

    /// <summary>
    /// Validates a saved package against the given Office version, returning the errors as text.
    /// </summary>
    /// <remarks>
    /// The errors are rendered while the package is still open: <c>ValidationErrorInfo</c> resolves
    /// its part and node lazily, so reading them after disposal throws.
    /// </remarks>
    public static IReadOnlyList<string> Validate(
        byte[] package, FileFormatVersions version = FileFormatVersions.Office2019)
    {
        using var stream = new MemoryStream(package);
        using var opened = WordprocessingDocument.Open(stream, false);

        return new OpenXmlValidator(version)
            .Validate(opened)
            .Select(error => $"[{error.Part?.Uri?.ToString() ?? "package"}] {error.Description}")
            .ToList();
    }

    /// <summary>A saved package held open for inspection.</summary>
    internal sealed class SavedPackage : IDisposable
    {
        private readonly MemoryStream _stream;

        public SavedPackage(byte[] package)
        {
            _stream = new MemoryStream(package);
            Document = WordprocessingDocument.Open(_stream, false);
            MainPart = Document.MainDocumentPart!;
        }

        public WordprocessingDocument Document { get; }
        public MainDocumentPart MainPart { get; }
        public Body Body => MainPart.Document.Body!;

        public void Dispose()
        {
            Document.Dispose();
            _stream.Dispose();
        }
    }
}
