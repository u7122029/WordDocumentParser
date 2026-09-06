using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;
using WordRun = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace WordDocumentParser.Tests;

/// <summary>
/// Regressions for the pre-scan's budget and media-type handling, and for field reconciliation.
/// </summary>
public class FourthReviewTests
{
    /// <summary>Builds the run sequence for a field with the given code and cached result.</summary>
    private static string Field(string code, string result) =>
        "<w:r><w:fldChar w:fldCharType='begin'/></w:r>" +
        $"<w:r><w:instrText> {code} </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType='separate'/></w:r>" +
        $"<w:r><w:t>{result}</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType='end'/></w:r>";

    // ---- Pre-scan --------------------------------------------------------

    [Fact]
    public void TheDepthScanStopsAtTheCharacterBudget()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var custom = p.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var stream = custom.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream);
            writer.Write($"<root attr='{new string('x', 4_000_000)}'/>");
        });

        static long AllocationsToReject(byte[] bytes, int depthLimit)
        {
            using var parser = new WordDocumentTreeParser
            {
                Limits = new DocumentLimits { MaxCharactersInPart = 4096, MaxElementDepth = depthLimit }
            };

            using var input = new MemoryStream(bytes);
            var before = GC.GetAllocatedBytesForCurrentThread();

            Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
            return GC.GetAllocatedBytesForCurrentThread() - before;
        }

        var withoutScan = AllocationsToReject(package, depthLimit: 0);
        var withScan = AllocationsToReject(package, depthLimit: 100);

        // Turning the depth scan on must not cost an extra copy of the oversized part: without a
        // budget on the scan's own reader it buffered the whole four-million-character attribute.
        Assert.True(withScan < withoutScan * 4 + 1_000_000,
            $"Scan allocated {withScan} bytes against {withoutScan} without it.");
    }

    [Fact]
    public void AnEmbeddedWorkbookIsNotScannedAsXml()
    {
        using var workbookBytes = new MemoryStream();

        using (var workbook = SpreadsheetDocument.Create(workbookBytes, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = workbook.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            workbookPart.Workbook.Append(new Sheets(new Sheet
            {
                Name = "Sheet1",
                SheetId = 1,
                Id = workbookPart.GetIdOfPart(worksheetPart)
            }));
        }

        var package = Create(Paragraph("Body"), p =>
        {
            var embedded = p.MainDocumentPart!.AddEmbeddedPackagePart(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            workbookBytes.Position = 0;
            embedded.FeedData(workbookBytes);
        });

        // The media type ends in ".sheet" but the part is a ZIP archive, not XML.
        using var parser = new WordDocumentTreeParser { Limits = DocumentLimits.Untrusted };
        using var input = new MemoryStream(package);

        Assert.Equal("Body", parser.ParseFromStream(input).Root.Children[0].Text);
    }

    // ---- Field reconciliation --------------------------------------------

    [Fact]
    public void RemovingOnlyTheFieldRunLeavesTheSurroundingTextIntact()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:t>Before </w:t></w:r>" +
            Field("DOCPROPERTY Title", "SECRET") +
            "<w:r><w:t> After</w:t></w:r></w:p>",
            p => p.PackageProperties.Title = "SECRET"));

        var node = document.Root.Children[0];
        Assert.Equal(1, node.Runs.RemoveAll(run => run.IsDocumentPropertyField));
        node.MarkRunsChanged();

        var saved = SavedBodyXml(document);
        Assert.Equal("Before  After", TextOf(saved));
        Assert.DoesNotContain("SECRET", saved, StringComparison.Ordinal);
        Assert.DoesNotContain("DOCPROPERTY", saved, StringComparison.Ordinal);
    }

    [Fact]
    public void ClearingRunsRemovesANonDocPropertyField()
    {
        var document = Parse(Create($"<w:p>{Field("MERGEFIELD Secret", "SECRET")}</w:p>"));

        var node = document.Root.Children[0];
        node.Runs.Clear();
        node.MarkRunsChanged();

        var saved = SavedBodyXml(document);
        Assert.Equal(string.Empty, TextOf(saved));
        Assert.DoesNotContain("MERGEFIELD", saved, StringComparison.Ordinal);
        Assert.DoesNotContain("fldChar", saved, StringComparison.Ordinal);
    }

    [Fact]
    public void EditingTextAroundAFieldKeepsTheField()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:t>Before </w:t></w:r>" +
            Field("DOCPROPERTY Title", "KEEP") +
            "<w:r><w:t> After</w:t></w:r></w:p>",
            p => p.PackageProperties.Title = "KEEP"));

        var node = document.Root.Children[0];
        node.Runs[0].Text = "Start ";

        var saved = SavedBodyXml(document);
        Assert.Contains("DOCPROPERTY", saved, StringComparison.Ordinal);
        Assert.Equal("Start KEEP After", TextOf(saved));
    }

    [Fact]
    public void ANonDocPropertyFieldSurvivesAnUnrelatedEdit()
    {
        var document = Parse(Create(
            $"<w:p>{Field("MERGEFIELD Name", "VALUE")}<w:r><w:t> tail</w:t></w:r></w:p>"));

        var node = document.Root.Children[0];
        node.Runs[^1].Text = " end";

        var saved = SavedBodyXml(document);
        Assert.Contains("MERGEFIELD", saved, StringComparison.Ordinal);
        Assert.Equal("VALUE end", TextOf(saved));
    }

    // ---- Formatting isolation on the fallback path -----------------------

    [Fact]
    public void ChangingTextAndFontTogetherLeavesRunSiblingsAlone()
    {
        var document = Parse(Create("<w:p><w:r><w:t>abc</w:t><w:tab/><w:t>def</w:t></w:r></w:p>"));

        var node = document.Root.Children[0];
        node.Runs[0].Text = "ABC";
        node.Runs[0].SetFont("Arial");

        using var saved = Open(Save(document));
        var runs = saved.MainPart.Document.Descendants<WordRun>()
            .Select(r => (r.InnerText, Font: r.RunProperties?.RunFonts?.Ascii?.Value))
            .Where(r => r.InnerText.Length > 0)
            .ToList();

        Assert.Contains(runs, r => r.InnerText == "ABC" && r.Font == "Arial");
        Assert.Contains(runs, r => r.InnerText == "def" && r.Font is null);
    }
}
