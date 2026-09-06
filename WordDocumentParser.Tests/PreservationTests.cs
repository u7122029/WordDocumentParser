using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;

namespace WordDocumentParser.Tests;

/// <summary>
/// Saving a document nobody edited must change nothing that matters.
/// </summary>
public class PreservationTests
{
    [Fact]
    public void UnchangedTableWithHyperlinkKeepsExactlyOneCopyOfItsText()
    {
        var package = Create(
            Table("<w:p><w:hyperlink r:id='rId1'><w:r><w:t>LINK</w:t></w:r></w:hyperlink></w:p>"),
            p => p.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rId1"));

        Assert.Equal("LINK", SavedText(Parse(package)));
    }

    [Fact]
    public void UnchangedCellKeepsEachRunsOwnFont()
    {
        var package = Create(Table(
            "<w:p><w:r><w:rPr><w:rFonts w:ascii='Arial'/></w:rPr><w:t>A</w:t></w:r>" +
            "<w:r><w:rPr><w:rFonts w:ascii='Courier New'/></w:rPr><w:t>B</w:t></w:r></w:p>"));

        var fonts = XDocument.Parse(SavedBodyXml(Parse(package)))
            .Descendants(XName.Get("rFonts", W))
            .Select(e => (string?)e.Attribute(XName.Get("ascii", W)))
            .ToList();

        Assert.Equal(["Arial", "Courier New"], fonts);
    }

    [Fact]
    public void UnchangedCheckboxKeepsItsExtensionNamespaceDefinition()
    {
        var package = Create(
            "<w:sdt><w:sdtPr><w:tag w:val='check'/>" +
            "<w14:checkbox xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'>" +
            "<w14:checked w14:val='1'/></w14:checkbox></w:sdtPr>" +
            $"<w:sdtContent>{Paragraph("Checked")}</w:sdtContent></w:sdt>");

        Assert.Contains("checkbox", SavedBodyXml(Parse(package)), StringComparison.Ordinal);
    }

    [Fact]
    public void CommentsPartSurvivesARoundTrip()
    {
        var package = Create(
            "<w:p><w:commentRangeStart w:id='0'/><w:r><w:t>Text</w:t></w:r>" +
            "<w:commentRangeEnd w:id='0'/><w:r><w:commentReference w:id='0'/></w:r></w:p>",
            p => p.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>().Comments =
                new Comments(new Comment(new Paragraph(new Run(new Text("Comment"))))
                {
                    Id = "0",
                    Author = "Reviewer"
                }));

        using var saved = Open(Save(Parse(package)));
        Assert.NotNull(saved.MainPart.WordprocessingCommentsPart);
    }

    [Fact]
    public void HeaderKeepsTheHyperlinkRelationshipItOwns()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var header = p.MainDocumentPart!.AddNewPart<HeaderPart>("rIdHead");
            header.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rIdLink");
            header.Header = new Header(
                $"<w:hdr xmlns:w='{W}' xmlns:r='{R}'><w:p><w:hyperlink r:id='rIdLink'>" +
                "<w:r><w:t>LINK</w:t></w:r></w:hyperlink></w:p></w:hdr>");
            p.MainDocumentPart.Document.Body!.Append(
                new SectionProperties(new HeaderReference { Id = "rIdHead", Type = HeaderFooterValues.Default }));
        });

        using var saved = Open(Save(Parse(package)));
        Assert.Single(saved.MainPart.HeaderParts.First().HyperlinkRelationships);
    }

    [Fact]
    public void CenteredParagraphKeepsItsAlignmentAndPageBreakThroughAFontChange()
    {
        var package = Create(
            "<w:p><w:pPr><w:jc w:val='center'/></w:pPr>" +
            "<w:r><w:t>Text</w:t><w:br w:type='page'/></w:r></w:p>");

        var document = Parse(package);
        Assert.Equal("center", document.Root.Children[0].ParagraphFormatting!.Alignment);

        document.Root.Children[0].SetParagraphFont("Arial");

        var saved = XDocument.Parse(SavedBodyXml(document));
        Assert.Equal("center", (string?)saved.Descendants(XName.Get("jc", W)).Single().Attribute(XName.Get("val", W)));
        Assert.Equal("page", (string?)saved.Descendants(XName.Get("br", W)).Single().Attribute(XName.Get("type", W)));
    }

    [Fact]
    public void ChangingAParagraphsFontKeepsItsHyperlink()
    {
        var package = Create(
            "<w:p><w:hyperlink r:id='rId1'><w:r><w:t>LINK</w:t></w:r></w:hyperlink></w:p>",
            p => p.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rId1"));

        var document = Parse(package);
        document.Root.Children[0].SetParagraphFont("Arial");

        var saved = XDocument.Parse(SavedBodyXml(document));
        Assert.Single(saved.Descendants(XName.Get("hyperlink", W)));
        Assert.Equal("LINK", TextOf(saved.ToString()));
        Assert.Contains(
            saved.Descendants(XName.Get("rFonts", W)),
            e => (string?)e.Attribute(XName.Get("ascii", W)) == "Arial");
    }

    [Fact]
    public void ReadingACustomPropertyDoesNotRetypeIt()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.AddCustomFilePropertiesPart();
            part.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(
                new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty(
                    new DocumentFormat.OpenXml.VariantTypes.VTInt32("42"))
                {
                    Name = "Count",
                    PropertyId = 2,
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                });
        });

        var document = Parse(package);
        Assert.Equal("42", document["Count"]);

        using var saved = Open(Save(document));
        Assert.Equal("i4", saved.Document.CustomFilePropertiesPart!.Properties.FirstChild!.FirstChild!.LocalName);
    }

    [Fact]
    public void ChangingACustomPropertyKeepsACompatibleType()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.AddCustomFilePropertiesPart();
            part.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(
                new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty(
                    new DocumentFormat.OpenXml.VariantTypes.VTInt32("42"))
                {
                    Name = "Count",
                    PropertyId = 2,
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                });
        });

        var document = Parse(package);
        document["Count"] = "43";

        using var saved = Open(Save(document));
        var property = saved.Document.CustomFilePropertiesPart!.Properties.FirstChild!;
        Assert.Equal("i4", property.FirstChild!.LocalName);
        Assert.Equal("43", property.FirstChild.InnerText);
    }

    [Fact]
    public void ChangingACustomPropertyToAnIncompatibleValueFallsBackToText()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.AddCustomFilePropertiesPart();
            part.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(
                new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty(
                    new DocumentFormat.OpenXml.VariantTypes.VTInt32("42"))
                {
                    Name = "Count",
                    PropertyId = 2,
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                });
        });

        var document = Parse(package);
        document["Count"] = "many";

        using var saved = Open(Save(document));
        var property = saved.Document.CustomFilePropertiesPart!.Properties.FirstChild!;
        Assert.Equal("lpwstr", property.FirstChild!.LocalName);
        Assert.Equal("many", property.FirstChild.InnerText);
    }

    [Fact]
    public void PreservationFailuresThrowByDefaultAndAreCollectedInRecoveryMode()
    {
        var document = new WordDocument();
        document.PackageData.HyperlinkRelationships["rId99"] =
            new Models.Package.HyperlinkRelationshipData { Url = "http://[not a uri", IsExternal = true };

        using var strictWriter = new WordDocumentTreeWriter();
        Assert.Throws<DocumentPreservationException>(() => strictWriter.BuildPackage(document));

        var recovery = new RecoveryOptions { ContinueOnPreservationFailure = true };
        using var lenientWriter = new WordDocumentTreeWriter { Recovery = recovery };
        lenientWriter.BuildPackage(document);

        Assert.True(recovery.HasDiagnostics);
        Assert.Equal("rId99", recovery.Diagnostics[0].Target);
    }
}
