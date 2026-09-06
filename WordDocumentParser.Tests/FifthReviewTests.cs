using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;

namespace WordDocumentParser.Tests;

/// <summary>
/// Regressions for matching surviving runs to the fields they came from, and for telling a
/// character-quota failure apart from malformed markup.
/// </summary>
public class FifthReviewTests
{
    private static string Field(string code, string result) =>
        "<w:r><w:fldChar w:fldCharType='begin'/></w:r>" +
        $"<w:r><w:instrText> {code} </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType='separate'/></w:r>" +
        $"<w:r><w:t>{result}</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType='end'/></w:r>";

    /// <summary>The field codes surviving in the saved document, in order.</summary>
    private static List<string> SavedFieldCodes(WordDocument document)
    {
        using var saved = Open(Save(document));
        return saved.MainPart.Document.Descendants<FieldCode>().Select(code => code.Text.Trim()).ToList();
    }

    /// <summary>A paragraph holding two property fields separated by plain text.</summary>
    private static byte[] TwoPropertyFields() => Create(
        "<w:p>" + Field("DOCPROPERTY Title", "SECRET") +
        "<w:r><w:t> / </w:t></w:r>" +
        Field("DOCPROPERTY Subject", "PUBLIC") + "</w:p>",
        p =>
        {
            p.PackageProperties.Title = "SECRET";
            p.PackageProperties.Subject = "PUBLIC";
        });

    [Fact]
    public void RemovingTheFirstOfTwoFieldsKeepsTheSecond()
    {
        var document = Parse(TwoPropertyFields());
        var node = document.Root.Children[0];

        Assert.Equal(1, node.Runs.RemoveAll(run => run.DocumentPropertyField?.PropertyName == "Title"));
        node.MarkRunsChanged();

        Assert.Equal(["DOCPROPERTY Subject"], SavedFieldCodes(document));
        Assert.DoesNotContain("SECRET", SavedBodyXml(document), StringComparison.Ordinal);
        Assert.Contains("PUBLIC", SavedBodyXml(document), StringComparison.Ordinal);
    }

    [Fact]
    public void RemovingTheSecondOfTwoFieldsKeepsTheFirst()
    {
        var document = Parse(TwoPropertyFields());
        var node = document.Root.Children[0];

        Assert.Equal(1, node.Runs.RemoveAll(run => run.DocumentPropertyField?.PropertyName == "Subject"));
        node.MarkRunsChanged();

        Assert.Equal(["DOCPROPERTY Title"], SavedFieldCodes(document));
        Assert.DoesNotContain("PUBLIC", SavedBodyXml(document), StringComparison.Ordinal);
        Assert.Contains("SECRET", SavedBodyXml(document), StringComparison.Ordinal);
    }

    [Fact]
    public void RemovingTheMiddleOfThreeFieldsKeepsTheOthers()
    {
        var document = Parse(Create(
            "<w:p>" + Field("DOCPROPERTY Title", "ONE") +
            Field("DOCPROPERTY Subject", "TWO") +
            Field("DOCPROPERTY Category", "THREE") + "</w:p>",
            p =>
            {
                p.PackageProperties.Title = "ONE";
                p.PackageProperties.Subject = "TWO";
                p.PackageProperties.Category = "THREE";
            }));

        var node = document.Root.Children[0];
        Assert.Equal(1, node.Runs.RemoveAll(run => run.DocumentPropertyField?.PropertyName == "Subject"));
        node.MarkRunsChanged();

        Assert.Equal(["DOCPROPERTY Title", "DOCPROPERTY Category"], SavedFieldCodes(document));
        Assert.DoesNotContain("TWO", SavedBodyXml(document), StringComparison.Ordinal);
    }

    [Fact]
    public void DeletingTextBeforeAFieldLeavesTheFieldAlone()
    {
        var document = Parse(Create(
            $"<w:p><w:r><w:t>Prefix </w:t></w:r>{Field("MERGEFIELD Name", "KEEP")}</w:p>"));

        var node = document.Root.Children[0];
        node.Runs.RemoveAt(0);
        node.MarkRunsChanged();

        var saved = SavedBodyXml(document);
        Assert.Contains("MERGEFIELD", saved, StringComparison.Ordinal);
        Assert.Equal("KEEP", TextOf(saved));
    }

    [Fact]
    public void DeletingTextAfterAFieldLeavesTheFieldAlone()
    {
        var document = Parse(Create(
            $"<w:p>{Field("MERGEFIELD Name", "KEEP")}<w:r><w:t> Suffix</w:t></w:r></w:p>"));

        var node = document.Root.Children[0];
        node.Runs.RemoveAt(node.Runs.Count - 1);
        node.MarkRunsChanged();

        var saved = SavedBodyXml(document);
        Assert.Contains("MERGEFIELD", saved, StringComparison.Ordinal);
        Assert.Equal("KEEP", TextOf(saved));
    }

    [Fact]
    public void DeletingAFieldResultBeforePlainTextRemovesItsInstruction()
    {
        var document = Parse(Create(
            $"<w:p>{Field("MERGEFIELD Secret", "SECRET")}<w:r><w:t> Tail</w:t></w:r></w:p>"));

        var node = document.Root.Children[0];
        node.Runs.RemoveAt(0);
        node.MarkRunsChanged();

        var saved = SavedBodyXml(document);
        Assert.Equal(" Tail", TextOf(saved));
        Assert.DoesNotContain("MERGEFIELD", saved, StringComparison.Ordinal);
        Assert.DoesNotContain("fldChar", saved, StringComparison.Ordinal);
    }

    [Fact]
    public void AFormattingChangeReachesASurvivingFieldAlongsideAnEditedNeighbour()
    {
        var document = Parse(Create(
            $"<w:p><w:r><w:t>Old </w:t></w:r>{Field("DOCPROPERTY Title", "KEEP")}</w:p>",
            p => p.PackageProperties.Title = "KEEP"));

        document.Root.Children[0].Runs[0].Text = "New ";
        document.SetDocumentFont("Arial");

        using var saved = Open(Save(document));
        var fonts = saved.MainPart.Document.Descendants<Run>()
            .Where(run => run.GetFirstChild<Text>() is not null)
            .Select(run => (run.GetFirstChild<Text>()!.Text, Font: run.RunProperties?.RunFonts?.Ascii?.Value))
            .ToList();

        Assert.Contains(fonts, entry => entry.Text == "New " && entry.Font == "Arial");
        Assert.Contains(fonts, entry => entry.Text == "KEEP" && entry.Font == "Arial");
    }

    [Fact]
    public void MalformedMultibyteXmlUnderTheBudgetReportsMalformedXml()
    {
        // 1,514 characters but 4,514 UTF-8 bytes: over a 4,096-byte count, under the character
        // budget. The failure is the mismatched end tag, not the quota.
        var malformed = $"<root>{new string('漢', 1500)}</wrong>";

        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var stream = part.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
            writer.Write(malformed);
        });

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxCharactersInPart = 4096, MaxElementDepth = 100 }
        };

        using var input = new MemoryStream(package);
        var failure = Assert.Throws<DocumentPreservationException>(() => parser.ParseFromStream(input));
        Assert.Contains("well-formed", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ValidXmlOverTheBudgetReportsTheLimit()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var stream = part.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
            writer.Write($"<root>{new string('x', 20_000)}</root>");
        });

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxCharactersInPart = 4096, MaxElementDepth = 100 }
        };

        using var input = new MemoryStream(package);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }
}
