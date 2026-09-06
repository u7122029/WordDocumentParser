using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using WordDocumentParser.Models.Formatting;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;

namespace WordDocumentParser.Tests;

/// <summary>Regressions for ordered source reconciliation, split formatting, and XML encodings.</summary>
public class SixthReviewTests
{
    private static string RunXml(string text) => $"<w:r><w:t>{text}</w:t></w:r>";
    private static string Field(string? result) =>
        "<w:r><w:fldChar w:fldCharType='begin'/></w:r>" +
        "<w:r><w:instrText> DOCPROPERTY Title </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType='separate'/></w:r>" +
        (result is null ? "" : RunXml(result)) +
        "<w:r><w:fldChar w:fldCharType='end'/></w:r>";

    [Theory]
    [InlineData(0, "XABC")]
    [InlineData(1, "AXBC")]
    [InlineData(2, "ABXC")]
    [InlineData(3, "ABCX")]
    public void InsertedRunsFollowTheirModelPosition(int position, string expected)
    {
        var document = Parse(Create($"<w:p>{RunXml("A")}{RunXml("B")}{RunXml("C")}</w:p>"));
        var node = document.Root.Children[0];
        node.Runs.Insert(position, new FormattedRun("X"));
        node.MarkRunsChanged();
        Assert.Equal(expected, SavedText(document));
        Assert.Equal(expected, SavedText(document));
        Assert.Empty(Validate(Save(document)));
    }

    [Theory]
    [InlineData(0, 2, 1)]
    [InlineData(1, 0, 2)]
    [InlineData(1, 2, 0)]
    [InlineData(2, 0, 1)]
    [InlineData(2, 1, 0)]
    public void ExistingRunsFollowTheirModelOrder(int first, int second, int third)
    {
        var document = Parse(Create($"<w:p>{RunXml("A")}{RunXml("B")}{RunXml("C")}</w:p>"));
        var node = document.Root.Children[0];
        var original = node.Runs.ToArray();
        node.Runs = [original[first], original[second], original[third]];
        Assert.Equal(string.Concat(node.Runs.Select(run => run.Text)), SavedText(document));
    }

    [Fact]
    public void ReorderingAndInsertingAroundAFieldKeepsTheWholeConstruct()
    {
        var document = Parse(Create($"<w:p>{RunXml("A")}{Field("KEEP")}{RunXml("B")}</w:p>"));
        var node = document.Root.Children[0];
        node.Runs = [node.Runs[2], new FormattedRun("X"), node.Runs[1], node.Runs[0]];
        using var saved = Open(Save(document));
        Assert.Equal("BXKEEPA", TextOf(saved.Body.OuterXml));
        Assert.Single(saved.Body.Descendants<FieldCode>());
        Assert.Equal([FieldCharValues.Begin, FieldCharValues.Separate, FieldCharValues.End],
            saved.Body.Descendants<FieldChar>().Select(field => field.FieldCharType!.Value));
        Assert.Empty(Validate(Save(document)));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("VALUE")]
    public void EditingAfterAPropertyFieldKeepsTheEditedText(string? result)
    {
        var source = Create($"<w:p>{Field(result)}{RunXml("OLD")}</w:p>");
        Assert.Empty(Validate(source));
        var document = Parse(source);
        document.Root.Children[0].Runs[^1].Text = "NEW";
        document.Root.Children[0].Runs[^1].SetFont("Arial");
        using var saved = Open(Save(document));
        Assert.Equal((result ?? "") + "NEW", TextOf(saved.Body.OuterXml));
        Assert.Single(saved.Body.Descendants<FieldCode>());
        Assert.Equal("Arial", saved.Body.Descendants<Run>().Last().RunProperties?.RunFonts?.Ascii?.Value);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("VALUE")]
    public void DeletingAPropertyFieldRemovesEvenAnAbsentCachedResult(string? result)
    {
        var document = Parse(Create($"<w:p>{Field(result)}{RunXml("KEEP")}</w:p>"));
        var node = document.Root.Children[0];
        node.Runs.RemoveAt(0);
        node.MarkRunsChanged();
        var xml = SavedBodyXml(document);
        Assert.Equal("KEEP", TextOf(xml));
        Assert.DoesNotContain("fldChar", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("DOCPROPERTY", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void PartialFontsStaySeparateWhenAnotherRunChanges()
    {
        var document = Parse(Create($"<w:p>{RunXml("abcdef")}{RunXml("old")}</w:p>"));
        var node = document.Root.Children[0];
        node.SetFontForText("bc", "Arial");
        node.SetFontForText("ef", "Courier New");
        node.Runs[^1].Text = "new";
        using var saved = Open(Save(document));
        var runs = saved.Body.Descendants<Run>().Select(run =>
            (run.InnerText, run.RunProperties?.RunFonts?.Ascii?.Value)).ToArray();
        Assert.Equal([("a", null), ("bc", "Arial"), ("d", null), ("ef", "Courier New"), ("new", null)], runs);
        Assert.Equal(saved.Body.OuterXml, SavedBodyXml(document));
    }

    [Fact]
    public void ReorderingWithinAControlPreservesItsPropertiesAndBookmarks()
    {
        var document = Parse(Create("<w:p><w:bookmarkStart w:id='0' w:name='mark'/>" +
            "<w:sdt><w:sdtPr><w:id w:val='7'/><w:tag w:val='control'/></w:sdtPr><w:sdtContent>" +
            RunXml("A") + RunXml("B") + "</w:sdtContent></w:sdt><w:bookmarkEnd w:id='0'/></w:p>"));
        var node = document.Root.Children[0];
        node.Runs.Reverse();
        node.Runs.Insert(1, new FormattedRun("X"));
        node.MarkRunsChanged();
        using var saved = Open(Save(document));
        var control = Assert.Single(saved.Body.Descendants<SdtRun>());
        Assert.Equal("BXA", control.SdtContentRun!.InnerText);
        Assert.Equal(7, control.SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value);
        Assert.Single(saved.Body.Descendants<BookmarkStart>());
        Assert.Single(saved.Body.Descendants<BookmarkEnd>());
        Assert.Empty(Validate(Save(document)));
    }

    [Fact]
    public void AContainerMovesWithItsRuns()
    {
        var document = Parse(Create("<w:p>" + RunXml("A") +
            "<w:hyperlink w:anchor='target'>" + RunXml("B") + RunXml("C") + "</w:hyperlink></w:p>"));
        var node = document.Root.Children[0];
        node.Runs = [node.Runs[1], node.Runs[2], node.Runs[0]];
        using var saved = Open(Save(document));
        Assert.Equal("BCA", TextOf(saved.Body.OuterXml));
        Assert.Equal("target", Assert.Single(saved.Body.Descendants<Hyperlink>()).Anchor?.Value);
    }

    [Fact]
    public void InterleavedContainersFailExplicitly()
    {
        var document = Parse(Create("<w:p>" + RunXml("A") +
            "<w:hyperlink w:anchor='target'>" + RunXml("B") + RunXml("C") + "</w:hyperlink></w:p>"));
        var node = document.Root.Children[0];
        node.Runs = [node.Runs[1], node.Runs[0], node.Runs[2]];
        Assert.Throws<DocumentPreservationException>(() => Save(document));
    }

    [Theory]
    [InlineData("utf-8", false)]
    [InlineData("utf-8", true)]
    [InlineData("utf-16LE", false)]
    [InlineData("utf-16BE", false)]
    [InlineData("utf-16", true)]
    [InlineData("utf-16BE", true)]
    [InlineData("iso-8859-1", false)]
    public void DepthScanHonoursXmlEncodings(string name, bool bom)
    {
        var bytes = EncodedPart(name, bom, "<root>café</root>");
        using var parser = new WordDocumentTreeParser { Limits = DocumentLimits.Untrusted };
        using var input = new MemoryStream(bytes);
        Assert.Equal("Body", parser.ParseFromStream(input).Root.Children[0].Text);
    }

    [Theory]
    [InlineData("utf-8")]
    [InlineData("utf-16LE")]
    [InlineData("utf-16BE")]
    public void EncodedXmlExceedingTheCharacterBudgetReportsTheLimit(string encoding)
    {
        var bytes = EncodedPart(encoding, false, $"<root attr='{new string('x', 20_000)}'/>");
        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxCharactersInPart = 4096, MaxElementDepth = 100 }
        };
        using var input = new MemoryStream(bytes);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

    [Theory]
    [InlineData("utf-16LE")]
    [InlineData("utf-16BE")]
    public void MalformedEncodedXmlBelowQuotaReportsMalformedXml(string encoding)
    {
        var bytes = EncodedPart(encoding, false, $"<root>{new string('漢', 1500)}</wrong>");
        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxCharactersInPart = 4096, MaxElementDepth = 100 }
        };
        using var input = new MemoryStream(bytes);
        Assert.Throws<DocumentPreservationException>(() => parser.ParseFromStream(input));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidUtf8ReportsMalformedXml(bool bom)
    {
        var bytes = Create(Paragraph("Body"), package =>
        {
            var part = package.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var stream = part.GetStream(FileMode.Create);
            if (bom) stream.Write(Encoding.UTF8.GetPreamble());
            stream.Write("<root>"u8);
            stream.WriteByte(0xFF);
            stream.Write("</root>"u8);
        });
        using var parser = new WordDocumentTreeParser { Limits = DocumentLimits.Untrusted };
        using var input = new MemoryStream(bytes);
        Assert.Throws<DocumentPreservationException>(() => parser.ParseFromStream(input));
    }

    [Theory]
    [InlineData("utf-16LE")]
    [InlineData("utf-16BE")]
    public void EncodedXmlStillEnforcesDepth(string encoding)
    {
        var xml = string.Concat(Enumerable.Repeat("<root>", 120)) +
                  string.Concat(Enumerable.Repeat("</root>", 120));
        var bytes = EncodedPart(encoding, false, xml);
        using var parser = new WordDocumentTreeParser { Limits = DocumentLimits.Untrusted };
        using var input = new MemoryStream(bytes);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

    [Fact]
    public void AnOversizedXmlDeclarationStopsAtTheCharacterBudget()
    {
        var bytes = Create(Paragraph("Body"), package =>
        {
            var part = package.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var stream = part.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream);
            writer.Write($"<?xml {new string(' ', 20_000)}version='1.0'?><root/>");
        });
        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxCharactersInPart = 4096, MaxElementDepth = 100 }
        };
        using var input = new MemoryStream(bytes);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

    [Fact]
    public void InsertedTabsAndBreaksRetainTheirKindsAndOrder()
    {
        var document = Parse(Create(Paragraph("A")));
        var node = document.Root.Children[0];
        node.Runs.Insert(0, new FormattedRun { IsTab = true });
        node.Runs.Add(new FormattedRun { IsBreak = true, BreakType = "page" });
        node.MarkRunsChanged();
        using var saved = Open(Save(document));
        Assert.Equal(["tab", "t", "br"], saved.Body.Descendants<Run>()
            .SelectMany(run => run.ChildElements.Where(child => child is not RunProperties))
            .Select(child => child.LocalName));
        Assert.Equal(BreakValues.Page, saved.Body.Descendants<Break>().Single().Type!.Value);
        Assert.Empty(Validate(Save(document)));
    }

    [Fact]
    public void ReorderingEqualTextFieldsStillUsesTheirIdentity()
    {
        var subject = Field("SAME").Replace("Title", "Subject", StringComparison.Ordinal);
        var document = Parse(Create($"<w:p>{Field("SAME")}{subject}</w:p>"));
        var node = document.Root.Children[0];
        node.Runs.Reverse();
        node.MarkRunsChanged();
        using var saved = Open(Save(document));
        Assert.Equal(["DOCPROPERTY Subject", "DOCPROPERTY Title"],
            saved.Body.Descendants<FieldCode>().Select(code => code.Text.Trim()));
    }

    [Theory]
    [InlineData(false, "BA")]
    [InlineData(true, "B")]
    public void NonCollapsedFieldResultsFollowModelEdits(bool deleteFirst, string expected)
    {
        var field = Field("A").Replace("DOCPROPERTY Title", "MERGEFIELD Name", StringComparison.Ordinal)
            .Replace(RunXml("A"), RunXml("A") + RunXml("B"), StringComparison.Ordinal);
        var document = Parse(Create($"<w:p>{field}</w:p>"));
        var node = document.Root.Children[0];
        if (deleteFirst) node.Runs.RemoveAt(0);
        else node.Runs.Reverse();
        node.MarkRunsChanged();
        using var saved = Open(Save(document));
        Assert.Equal(expected, TextOf(saved.Body.OuterXml));
        Assert.Equal("MERGEFIELD Name", Assert.Single(saved.Body.Descendants<FieldCode>()).Text.Trim());
        Assert.Equal([FieldCharValues.Begin, FieldCharValues.Separate, FieldCharValues.End],
            saved.Body.Descendants<FieldChar>().Select(fieldChar => fieldChar.FieldCharType!.Value));
        Assert.Empty(Validate(Save(document)));
    }

    private static byte[] EncodedPart(string name, bool bom, string content) => Create(Paragraph("Body"), package =>
    {
        var encoding = Encoding.GetEncoding(name);
        var part = package.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        using var stream = part.GetStream(FileMode.Create);
        if (bom) stream.Write(encoding.GetPreamble());
        stream.Write(encoding.GetBytes($"<?xml version='1.0' encoding='{name}'?>{content}"));
    });
}
