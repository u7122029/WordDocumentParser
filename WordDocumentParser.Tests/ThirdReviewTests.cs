
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;

namespace WordDocumentParser.Tests;

/// <summary>
/// Regressions for field deletion, formatting isolation, control metadata, and hostile nesting.
/// </summary>
public class ThirdReviewTests
{
    /// <summary>A paragraph whose text comes from a DOCPROPERTY field.</summary>
    private const string DocPropertyFieldParagraph =
        "<w:p><w:r><w:fldChar w:fldCharType='begin'/></w:r>" +
        "<w:r><w:instrText> DOCPROPERTY Title </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType='separate'/></w:r>" +
        "<w:r><w:t>SECRET</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType='end'/></w:r></w:p>";

    [Fact]
    public void ClearingRunsRemovesFieldTextAsWellAsPlainText()
    {
        var document = Parse(Create(DocPropertyFieldParagraph, p => p.PackageProperties.Title = "SECRET"));

        document.Root.Children[0].Runs.Clear();
        document.Root.Children[0].MarkRunsChanged();

        var saved = SavedBodyXml(document);
        Assert.Equal(string.Empty, TextOf(saved));

        // The field construct must be gone, not merely emptied: a surviving field recomputes itself.
        Assert.DoesNotContain("DOCPROPERTY", saved, StringComparison.Ordinal);
        Assert.DoesNotContain("fldChar", saved, StringComparison.Ordinal);
    }

    [Fact]
    public void AFieldParagraphIsUntouchedWhenNothingIsEdited()
    {
        var document = Parse(Create(DocPropertyFieldParagraph, p => p.PackageProperties.Title = "SECRET"));

        var saved = SavedBodyXml(document);
        Assert.Contains("DOCPROPERTY", saved, StringComparison.Ordinal);
        Assert.Equal("SECRET", TextOf(saved));
    }

    [Fact]
    public void StylingOneTextElementLeavesItsRunSiblingsAlone()
    {
        var document = Parse(Create("<w:p><w:r><w:t>abc</w:t><w:tab/><w:t>def</w:t></w:r></w:p>"));

        Assert.Equal(1, document.Root.Children[0].SetFontForText("abc", "Arial"));

        var runs = XDocument.Parse(SavedBodyXml(document))
            .Descendants(XName.Get("r", W))
            .Select(r => (
                Text: string.Concat(r.Descendants(XName.Get("t", W)).Select(t => t.Value)),
                Font: (string?)r.Descendants(XName.Get("rFonts", W))
                    .Select(f => f.Attribute(XName.Get("ascii", W))).FirstOrDefault()))
            .ToList();

        Assert.Contains(runs, r => r.Text == "abc" && r.Font == "Arial");
        Assert.Contains(runs, r => r.Text == "def" && r.Font is null);
    }

    [Fact]
    public void AFontChangeKeepsEachInlineControlRunsOwnFormatting()
    {
        var document = Parse(Create(
            "<w:p><w:sdt><w:sdtPr><w:id w:val='7'/><w:tag w:val='name'/></w:sdtPr><w:sdtContent>" +
            "<w:r><w:rPr><w:b/></w:rPr><w:t>AAA</w:t></w:r>" +
            "<w:r><w:rPr><w:i/></w:rPr><w:t>BBB</w:t></w:r>" +
            "</w:sdtContent></w:sdt></w:p>"));

        document.SetDocumentFont("Arial");

        using var saved = Open(Save(document));
        var runs = saved.MainPart.Document.Descendants<Run>()
            .Where(r => r.InnerText.Length > 0)
            .Select(r => (r.InnerText, Bold: r.RunProperties?.Bold is not null, Italic: r.RunProperties?.Italic is not null))
            .ToList();

        Assert.Equal([("AAA", true, false), ("BBB", false, true)], runs);
    }

    [Fact]
    public void StylingAControlsTextKeepsItReachableThroughItsControl()
    {
        var document = Parse(Create(
            "<w:p><w:sdt><w:sdtPr><w:id w:val='7'/><w:tag w:val='name'/></w:sdtPr>" +
            "<w:sdtContent><w:r><w:t>OLD</w:t></w:r></w:sdtContent></w:sdt></w:p>"));

        document.Root.Children[0].SetFontForText("OLD", "Arial");

        Assert.True(document.SetContentControlValueByTag("name", "NEW"));
        Assert.Equal("NEW", SavedText(document));
    }

    [Fact]
    public void StylingPartOfAControlsTextKeepsTheRestReachable()
    {
        var document = Parse(Create(
            "<w:p><w:sdt><w:sdtPr><w:id w:val='7'/><w:tag w:val='name'/></w:sdtPr>" +
            "<w:sdtContent><w:r><w:t>OLDVALUE</w:t></w:r></w:sdtContent></w:sdt></w:p>"));

        document.Root.Children[0].SetFontForText("VALUE", "Arial");

        Assert.True(document.SetContentControlValueByTag("name", "NEW"));
        Assert.Equal("NEW", SavedText(document));
    }

    [Fact]
    public void RemovingAContentControlAroundAHeadingKeepsTheSectionThatFollows()
    {
        var document = Parse(Create(
            "<w:sdt><w:sdtPr><w:tag w:val='name'/></w:sdtPr><w:sdtContent>" +
            Paragraph("Heading", 1) + "</w:sdtContent></w:sdt>" +
            Paragraph("One") + Paragraph("Two")));

        Assert.True(document.RemoveContentControlByTag("name"));

        var saved = SavedBodyXml(document);
        Assert.Equal("HeadingOneTwo", TextOf(saved));
        Assert.DoesNotContain("<w:sdt", saved, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("decimal", "1.5", "1,234")]
    [InlineData("filetime", "2026-01-01T00:00:00Z", "09/06/2026")]
    public void AValueOutsideTheTypesXmlFormFallsBackToText(string variant, string original, string replacement)
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.AddCustomFilePropertiesPart();
            using var stream = part.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream);
            writer.Write(
                "<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties' " +
                "xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'>" +
                "<property fmtid='{D5CDD505-2E9C-101B-9397-08002B2CF9AE}' pid='2' name='Value'>" +
                $"<vt:{variant}>{original}</vt:{variant}></property></Properties>");
        });

        var document = Parse(package);
        document["Value"] = replacement;

        var bytes = Save(document);

        using var saved = Open(bytes);
        var property = saved.Document.CustomFilePropertiesPart!.Properties.FirstChild!;
        Assert.Equal("lpwstr", property.FirstChild!.LocalName);

        Assert.Empty(Validate(bytes));
    }

    [Theory]
    [InlineData("decimal", "1.5", "2.5")]
    [InlineData("filetime", "2026-01-01T00:00:00Z", "2027-02-03T04:05:06Z")]
    public void AValueMatchingTheTypesXmlFormKeepsThatType(string variant, string original, string replacement)
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.AddCustomFilePropertiesPart();
            using var stream = part.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream);
            writer.Write(
                "<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties' " +
                "xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'>" +
                "<property fmtid='{D5CDD505-2E9C-101B-9397-08002B2CF9AE}' pid='2' name='Value'>" +
                $"<vt:{variant}>{original}</vt:{variant}></property></Properties>");
        });

        var document = Parse(package);
        document["Value"] = replacement;

        var bytes = Save(document);

        using var saved = Open(bytes);
        var property = saved.Document.CustomFilePropertiesPart!.Properties.FirstChild!;
        Assert.Equal(variant, property.FirstChild!.LocalName);

        Assert.Empty(Validate(bytes));
    }

    [Fact]
    public void DeeplyNestedXmlIsRejectedRatherThanOverflowingTheStack()
    {
        // Ten thousand nested tags in about two kilobytes. The SDK loads its object model by
        // recursive descent, so this must be rejected before any DOM is built. A stack overflow is
        // not catchable, so a regression here does not fail this assertion — it takes the whole test
        // host down, and the run reports a crash instead of a failure.
        using var buffer = new MemoryStream();
        buffer.Write(Create(Paragraph("Body")));

        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            zip.GetEntry("word/document.xml")!.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open());

            writer.Write($"<w:document xmlns:w='{W}'><w:body>");
            for (var i = 0; i < 10_000; i++) writer.Write("<w:sdt><w:sdtContent>");
            writer.Write(Paragraph("Deep"));
            for (var i = 0; i < 10_000; i++) writer.Write("</w:sdtContent></w:sdt>");
            writer.Write("</w:body></w:document>");
        }

        buffer.Position = 0;

        using var parser = new WordDocumentTreeParser { Limits = DocumentLimits.Untrusted };
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(buffer));
    }

    [Fact]
    public void DeeplyNestedContentControlsAreRejected()
    {
        var xml = Paragraph("Deep");
        for (var i = 0; i < 12; i++)
        {
            xml = $"<w:sdt><w:sdtPr/><w:sdtContent>{xml}</w:sdtContent></w:sdt>";
        }

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxElementDepth = 8 }
        };

        using var input = new MemoryStream(Create(xml));
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

}
