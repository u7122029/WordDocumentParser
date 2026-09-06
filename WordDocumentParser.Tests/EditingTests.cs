using System.Xml.Linq;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;

namespace WordDocumentParser.Tests;

/// <summary>
/// Edits made through the model must reach the saved document, and must not reach anything else.
/// </summary>
public class EditingTests
{
    [Fact]
    public void AssigningNodeTextPersists()
    {
        var document = Parse(Create(Paragraph("OLD")));
        document.Root.Children[0].Text = "NEW";

        Assert.Equal("NEW", SavedText(document));
    }

    [Fact]
    public void SettingABlockContentControlValuePersists()
    {
        var document = Parse(Create(
            "<w:sdt><w:sdtPr><w:tag w:val='name'/><w:text/></w:sdtPr>" +
            $"<w:sdtContent>{Paragraph("OLD")}</w:sdtContent></w:sdt>"));

        Assert.True(document.SetContentControlValueByTag("name", "NEW"));
        Assert.Equal("NEW", SavedText(document));
    }

    [Fact]
    public void SettingAnInlineContentControlValueLeavesSurroundingTextAlone()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:t>Before </w:t></w:r>" +
            "<w:sdt><w:sdtPr><w:id w:val='7'/><w:tag w:val='name'/><w:text/></w:sdtPr>" +
            "<w:sdtContent><w:r><w:t>OLD</w:t></w:r></w:sdtContent></w:sdt>" +
            "<w:r><w:t> After</w:t></w:r></w:p>"));

        Assert.True(document.SetContentControlValueByTag("name", "NEW"));
        Assert.Equal("Before NEW After", document.Root.Children[0].GetText());
        Assert.Equal("Before NEW After", SavedText(document));
    }

    [Fact]
    public void ClearingACellRemovesItsTextFromTheSavedDocument()
    {
        var document = Parse(Create(Table(Paragraph("SECRET"))));
        document.FindAllTables().First().GetCell(0, 0)!.ClearContent();

        Assert.Equal(string.Empty, SavedText(document));
    }

    [Fact]
    public void SettingACellToEmptyRemovesItsTextFromTheSavedDocument()
    {
        var document = Parse(Create(Table(Paragraph("SECRET"))));
        Assert.True(document.FindAllTables().First().SetCellText(0, 0, ""));

        Assert.Equal(string.Empty, SavedText(document));
    }

    [Fact]
    public void AddingARowToANestedTablePersists()
    {
        var document = Parse(Create(Table(Table(Paragraph("OLD")) + Paragraph(""))));
        document.FindAllTables().Skip(1).First().AddRow("NEW");

        Assert.Contains("NEW", SavedText(document), StringComparison.Ordinal);
    }

    [Fact]
    public void SettingTableAlignmentTargetsTheTableNotItsCellParagraphs()
    {
        var document = Parse(Create(Table(
            "<w:p><w:pPr><w:jc w:val='right'/></w:pPr><w:r><w:t>Cell</w:t></w:r></w:p>")));

        document.FindAllTables().First().SetTableAlignment("Center");

        var saved = XDocument.Parse(SavedBodyXml(document));

        var tableAlignment = saved.Descendants(XName.Get("tblPr", W))
            .Elements(XName.Get("jc", W))
            .Select(e => (string?)e.Attribute(XName.Get("val", W)))
            .SingleOrDefault();
        Assert.Equal("center", tableAlignment);

        var paragraphAlignment = saved.Descendants(XName.Get("pPr", W))
            .Elements(XName.Get("jc", W))
            .Select(e => (string?)e.Attribute(XName.Get("val", W)))
            .Single();
        Assert.Equal("right", paragraphAlignment);
    }

    [Fact]
    public void SettingCellBordersPersists()
    {
        var document = Parse(Create(Table(Paragraph("Cell"))));
        document.FindAllTables().First().GetCell(0, 0)!.SetBorders();

        Assert.Single(XDocument.Parse(SavedBodyXml(document)).Descendants(XName.Get("tcBorders", W)));
    }

    [Fact]
    public void ChangingStyleOnAParagraphWithEmptyPropertiesProducesAValidDocument()
    {
        var document = Parse(Create("<w:p><w:pPr/><w:r><w:t>Text</w:t></w:r></w:p>"));
        document.Root.Children[0].ChangeStyle("Heading2");

        Assert.Empty(Validate(Save(document)));
        Assert.Contains("Heading2", SavedBodyXml(document), StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentWideFontChangeReachesTextInsideTables()
    {
        var document = Parse(Create(Table(Paragraph("Cell"))));

        Assert.Equal(1, document.SetDocumentFont("Arial"));
        Assert.Contains(
            XDocument.Parse(SavedBodyXml(document)).Descendants(XName.Get("rFonts", W)),
            e => (string?)e.Attribute(XName.Get("ascii", W)) == "Arial");
    }

    [Fact]
    public void SetFontForTextTerminatesWhenChangingEveryOccurrence()
    {
        var node = new DocumentNode(ContentType.Paragraph, "match and match again");

        var changed = TimeBoxed(() => node.SetFontForText("match", "Arial", allOccurrences: true));

        Assert.Equal(2, changed);
        Assert.Equal("match and match again", string.Concat(node.Runs.Select(r => r.Text)));
        Assert.Equal(2, node.Runs.Count(r => r.GetFont() == "Arial"));
    }

    [Fact]
    public void SetFontForTextChangesOnlyTheMatchedSpan()
    {
        var node = new DocumentNode(ContentType.Paragraph, "keep match keep");

        Assert.Equal(1, node.SetFontForText("match", "Arial"));

        var styled = node.Runs.Where(r => r.GetFont() == "Arial").Select(r => r.Text).ToList();
        Assert.Equal(["match"], styled);
    }

    [Fact]
    public void APartialSpanFontChangeSplitsTheRunInTheSavedDocument()
    {
        var document = Parse(Create("<w:p><w:r><w:t>keep match keep</w:t></w:r></w:p>"));

        Assert.Equal(1, document.Root.Children[0].SetFontForText("match", "Arial"));

        var saved = XDocument.Parse(SavedBodyXml(document));

        // The text must survive intact...
        Assert.Equal("keep match keep", TextOf(saved.ToString()));

        // ...and only the matched span carries the new font.
        var styledText = saved.Descendants(XName.Get("r", W))
            .Where(run => run.Descendants(XName.Get("rFonts", W))
                .Any(f => (string?)f.Attribute(XName.Get("ascii", W)) == "Arial"))
            .Select(run => string.Concat(run.Descendants(XName.Get("t", W)).Select(t => t.Value)));

        Assert.Equal(["match"], styledText);
    }

    [Fact]
    public void APartialSpanFontChangeAcrossExistingRunsStylesOnlyTheMatch()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:t>ab</w:t></w:r><w:r><w:t>cd</w:t></w:r><w:r><w:t>ef</w:t></w:r></w:p>"));

        Assert.Equal(1, document.Root.Children[0].SetFontForText("bcde", "Arial"));

        var saved = XDocument.Parse(SavedBodyXml(document));
        Assert.Equal("abcdef", TextOf(saved.ToString()));

        var styledText = string.Concat(saved.Descendants(XName.Get("r", W))
            .Where(run => run.Descendants(XName.Get("rFonts", W))
                .Any(f => (string?)f.Attribute(XName.Get("ascii", W)) == "Arial"))
            .Select(run => string.Concat(run.Descendants(XName.Get("t", W)).Select(t => t.Value))));

        Assert.Equal("bcde", styledText);
    }

    [Fact]
    public void SavingTwiceProducesTheSameDocument()
    {
        var package = Create(
            Paragraph("Title", 1) +
            "<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>" +
            Table(Paragraph("Cell")));

        var once = Save(Parse(package));
        var twice = Save(Parse(once));

        Assert.Equal(TextOf(SavedBodyXml(Parse(once))), TextOf(SavedBodyXml(Parse(twice))));
        Assert.Equal(SavedBodyXml(Parse(once)), SavedBodyXml(Parse(twice)));
    }

    [Fact]
    public void RemovingAColumnInsideAMergedSpanKeepsTheFollowingCell()
    {
        var document = Parse(Create(
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:gridSpan w:val='2'/></w:tcPr>" + Paragraph("Merged") + "</w:tc>" +
            "<w:tc>" + Paragraph("Keep") + "</w:tc></w:tr></w:tbl>"));

        var table = document.FindAllTables().First();
        Assert.True(table.RemoveColumn(1));

        Assert.Equal(["Merged", "Keep"], table.GetRowCells(0).Select(c => c.TextContent));

        var savedText = SavedText(document);
        Assert.Contains("Merged", savedText, StringComparison.Ordinal);
        Assert.Contains("Keep", savedText, StringComparison.Ordinal);
    }

    /// <summary>
    /// Runs an operation on a worker thread and fails the test if it does not finish promptly, so a
    /// regression to an unbounded loop reports as a failure instead of hanging the test run.
    /// </summary>
    private static T TimeBoxed<T>(Func<T> operation, int timeoutMilliseconds = 5000)
    {
        T result = default!;
        Exception? failure = null;

        var thread = new Thread(() =>
        {
            try { result = operation(); }
            catch (Exception ex) { failure = ex; }
        }) { IsBackground = true };

        thread.Start();
        Assert.True(thread.Join(timeoutMilliseconds), "Operation did not terminate.");

        if (failure is not null) throw failure;
        return result;
    }
}
