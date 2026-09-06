using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for changing fonts on runs, text spans, and paragraphs.
/// Note: These methods change the actual font family applied to text, not the paragraph style ID.
/// </summary>
/// <remarks>
/// Font changes are recorded on the runs they apply to. The writer then edits the font of exactly
/// those runs in the paragraph's original XML, so hyperlinks, fields, bookmarks, and drawings around
/// the text survive a font change instead of being discarded with the rest of the paragraph.
/// </remarks>
public static class FontExtensions
{
    #region Run-level font changes

    /// <summary>
    /// Sets the font family for a formatted run.
    /// </summary>
    /// <param name="run">The run to modify</param>
    /// <param name="fontName">The font family name (e.g., "Calibri", "Arial", "Cascadia Code")</param>
    public static void SetFont(this FormattedRun run, string fontName)
    {
        run.Formatting.FontFamily = fontName;
        run.Formatting.FontFamilyAscii = fontName;
        // Also set for other character sets for consistency
        run.Formatting.FontFamilyEastAsia = fontName;
        run.Formatting.FontFamilyComplexScript = fontName;

        // Assigning the same font a run already had is not a change to the value, but it is still an
        // explicit instruction to write that font, so record it either way.
        MarkFontChanged(run);
    }

    /// <summary>
    /// Sets the font family for a formatted run with separate settings for different character sets.
    /// </summary>
    /// <param name="run">The run to modify</param>
    /// <param name="ascii">Font for ASCII characters</param>
    /// <param name="highAnsi">Font for high ANSI characters (optional, defaults to ascii)</param>
    /// <param name="eastAsia">Font for East Asian characters (optional)</param>
    /// <param name="complexScript">Font for complex scripts like Arabic/Hebrew (optional)</param>
    public static void SetFont(this FormattedRun run, string ascii, string? highAnsi = null, string? eastAsia = null, string? complexScript = null)
    {
        run.Formatting.FontFamilyAscii = ascii;
        run.Formatting.FontFamily = highAnsi ?? ascii;
        run.Formatting.FontFamilyEastAsia = eastAsia;
        run.Formatting.FontFamilyComplexScript = complexScript;

        MarkFontChanged(run);
    }

    private static void MarkFontChanged(FormattedRun run)
    {
        foreach (var property in RunFormattingFontProperties)
        {
            run.Formatting.MarkChanged(property);
        }
    }

    private static readonly string[] RunFormattingFontProperties =
    [
        nameof(RunFormatting.FontFamily), nameof(RunFormatting.FontFamilyAscii),
        nameof(RunFormatting.FontFamilyEastAsia), nameof(RunFormatting.FontFamilyComplexScript)
    ];

    /// <summary>
    /// Gets the font family name from a formatted run.
    /// </summary>
    /// <param name="run">The run to check</param>
    /// <returns>The font family name, or null if not set</returns>
    public static string? GetFont(this FormattedRun run)
        => run.Formatting.FontFamilyAscii ?? run.Formatting.FontFamily;

    #endregion

    #region Paragraph-level font changes

    /// <summary>
    /// Sets the font family for all runs in a paragraph node.
    /// This changes the actual font applied to all text, not the paragraph style.
    /// </summary>
    /// <param name="node">The paragraph node to modify</param>
    /// <param name="fontName">The font family name (e.g., "Calibri", "Arial")</param>
    /// <returns>True when the node was a paragraph-like node with text to restyle.</returns>
    public static bool SetParagraphFont(this DocumentNode node, string fontName)
    {
        if (node.Type is not (ContentType.Paragraph or ContentType.Heading or ContentType.ListItem))
            return false;

        var text = node.GetText();
        if (string.IsNullOrEmpty(text) && string.IsNullOrEmpty(node.Text))
            return false;

        if (node.HasFormattedRuns)
        {
            foreach (var run in node.Runs)
            {
                run.SetFont(fontName);
            }
        }
        else
        {
            var run = new FormattedRun(!string.IsNullOrEmpty(node.Text) ? node.Text : text);
            run.SetFont(fontName);
            node.Runs.Add(run);
            node.MarkRunsChanged();
        }

        return true;
    }

    /// <summary>
    /// Sets the font family for all paragraphs in a document, including text inside table cells.
    /// </summary>
    /// <param name="document">The document to modify</param>
    /// <param name="fontName">The font family name</param>
    /// <returns>The number of paragraphs modified</returns>
    public static int SetDocumentFont(this WordDocument document, string fontName)
        => document.Root.SetDocumentFont(fontName);

    /// <summary>
    /// Sets the font family for all paragraphs under a node, including text inside table cells.
    /// </summary>
    /// <param name="root">The root node to start from</param>
    /// <param name="fontName">The font family name</param>
    /// <returns>The number of paragraphs modified</returns>
    /// <remarks>
    /// Table cell content hangs off the table's model rather than off <c>Children</c>, so a traversal
    /// of the tree alone walks straight past every paragraph inside every table.
    /// </remarks>
    public static int SetDocumentFont(this DocumentNode root, string fontName)
    {
        var count = 0;
        foreach (var node in root.FindAllContent(n =>
                     n.Type is ContentType.Paragraph or ContentType.Heading or ContentType.ListItem))
        {
            if (node.SetParagraphFont(fontName))
            {
                count++;
            }
        }
        return count;
    }

    #endregion

    #region Span-level font changes (text range within a paragraph)

    /// <summary>
    /// Sets the font for a specific substring within a paragraph.
    /// If the text spans multiple runs, they will be split and the font applied to matching portions.
    /// </summary>
    /// <param name="node">The paragraph node to modify</param>
    /// <param name="searchText">The text to find and change font for</param>
    /// <param name="fontName">The font family name to apply</param>
    /// <param name="allOccurrences">If true, changes all occurrences; if false, only the first</param>
    /// <returns>The number of occurrences modified</returns>
    /// <remarks>
    /// Each match advances the search past the text just restyled. Restyling does not change the
    /// text, so a search that restarted from the beginning would re-find the same match forever.
    /// </remarks>
    public static int SetFontForText(this DocumentNode node, string searchText, string fontName, bool allOccurrences = false)
    {
        if (node.Type is not (ContentType.Paragraph or ContentType.Heading or ContentType.ListItem))
            return 0;

        if (string.IsNullOrEmpty(searchText))
            return 0;

        EnsureRuns(node);

        var fullText = string.Concat(node.Runs.Select(RunText));

        // Collect every match before touching the runs, then rebuild once. Splitting per match
        // rebuilt the whole (growing) run collection each time, so the cost grew with the square of
        // the number of matches.
        var ranges = new List<(int Start, int Length)>();
        var searchFrom = 0;

        while (searchFrom <= fullText.Length - searchText.Length)
        {
            var index = fullText.IndexOf(searchText, searchFrom, StringComparison.Ordinal);
            if (index < 0) break;

            ranges.Add((index, searchText.Length));
            if (!allOccurrences) break;

            searchFrom = index + searchText.Length;
        }

        if (ranges.Count == 0) return 0;

        ApplyFontToRanges(node, ranges, fontName);
        return ranges.Count;
    }

    /// <summary>
    /// Sets the font for a character range within a paragraph (by start index and length).
    /// </summary>
    /// <param name="node">The paragraph node to modify</param>
    /// <param name="startIndex">The starting character index (0-based)</param>
    /// <param name="length">The number of characters to apply the font to</param>
    /// <param name="fontName">The font family name to apply</param>
    /// <returns>True if the range was valid and font was applied</returns>
    public static bool SetFontForRange(this DocumentNode node, int startIndex, int length, string fontName)
    {
        if (node.Type is not (ContentType.Paragraph or ContentType.Heading or ContentType.ListItem))
            return false;

        if (startIndex < 0 || length <= 0)
            return false;

        EnsureRuns(node);

        var fullText = string.Concat(node.Runs.Select(RunText));
        if (startIndex + length > fullText.Length)
            return false;

        ApplyFontToRange(node, startIndex, length, fontName);
        return true;
    }

    /// <summary>
    /// Sets the font for runs matching a predicate.
    /// </summary>
    /// <param name="node">The paragraph node to modify</param>
    /// <param name="predicate">Condition to match runs</param>
    /// <param name="fontName">The font family name to apply</param>
    /// <returns>The number of runs modified</returns>
    public static int SetFontWhere(this DocumentNode node, Func<FormattedRun, bool> predicate, string fontName)
    {
        if (!node.HasFormattedRuns)
            return 0;

        var count = 0;
        foreach (var run in node.Runs.Where(predicate))
        {
            run.SetFont(fontName);
            count++;
        }

        return count;
    }

    #endregion

    #region Font queries

    /// <summary>
    /// Gets all unique font families used in a paragraph.
    /// </summary>
    /// <param name="node">The paragraph node to check</param>
    /// <returns>Set of font family names used</returns>
    public static HashSet<string> GetFontsUsed(this DocumentNode node)
    {
        var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var run in node.Runs)
        {
            var font = run.GetFont();
            if (!string.IsNullOrEmpty(font))
            {
                fonts.Add(font);
            }
        }

        return fonts;
    }

    /// <summary>
    /// Gets all unique font families used in a document, including text inside table cells.
    /// </summary>
    /// <param name="document">The document to analyze</param>
    /// <returns>Set of font family names used</returns>
    public static HashSet<string> GetAllFontsUsed(this WordDocument document)
        => document.Root.GetAllFontsUsed();

    /// <summary>
    /// Gets all unique font families used under a node, including table cell content.
    /// </summary>
    /// <param name="root">The root node to analyze</param>
    /// <returns>Set of font family names used</returns>
    public static HashSet<string> GetAllFontsUsed(this DocumentNode root)
    {
        var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var node in root.FindAllContent(_ => true))
        {
            fonts.UnionWith(node.GetFontsUsed());
        }

        return fonts;
    }

    /// <summary>
    /// Replaces one font with another throughout a document, including text inside table cells.
    /// </summary>
    /// <param name="document">The document to modify</param>
    /// <param name="fromFont">The font to replace</param>
    /// <param name="toFont">The font to replace with</param>
    /// <returns>The number of runs modified</returns>
    public static int ReplaceFont(this WordDocument document, string fromFont, string toFont)
        => document.Root.ReplaceFont(fromFont, toFont);

    /// <summary>
    /// Replaces one font with another under a node, including text inside table cells.
    /// </summary>
    /// <param name="root">The root node to start from</param>
    /// <param name="fromFont">The font to replace</param>
    /// <param name="toFont">The font to replace with</param>
    /// <returns>The number of runs modified</returns>
    public static int ReplaceFont(this DocumentNode root, string fromFont, string toFont)
    {
        var count = 0;

        foreach (var node in root.FindAllContent(n => n.HasFormattedRuns))
        {
            foreach (var run in node.Runs)
            {
                var currentFont = run.GetFont();
                if (currentFont is not null && currentFont.Equals(fromFont, StringComparison.OrdinalIgnoreCase))
                {
                    run.SetFont(toFont);
                    count++;
                }
            }
        }

        return count;
    }

    #endregion

    #region Private helpers

    private static string RunText(FormattedRun run) => run.IsTab ? "\t" : run.IsBreak ? " " : run.Text;

    private static void EnsureRuns(DocumentNode node)
    {
        if (node.HasFormattedRuns || string.IsNullOrEmpty(node.Text)) return;

        node.Runs.Add(new FormattedRun(node.Text));
        node.MarkRunsChanged();
    }

    /// <summary>
    /// Applies a font to a character range by splitting runs as needed.
    /// </summary>
    private static void ApplyFontToRange(DocumentNode node, int startIndex, int length, string fontName)
        => ApplyFontToRanges(node, [(startIndex, length)], fontName);

    /// <summary>
    /// Applies a font to several character ranges in one forward pass over the runs.
    /// </summary>
    /// <param name="node">The paragraph node to modify.</param>
    /// <param name="ranges">Non-overlapping ranges in ascending order.</param>
    /// <param name="fontName">The font to apply.</param>
    private static void ApplyFontToRanges(
        DocumentNode node, List<(int Start, int Length)> ranges, string fontName)
    {
        var newRuns = new List<FormattedRun>(node.Runs.Count + ranges.Count * 2);
        var position = 0;
        var nextRange = 0;

        foreach (var run in node.Runs)
        {
            var runText = RunText(run);
            var runStart = position;
            var runEnd = position + runText.Length;
            position = runEnd;

            // Ranges that end before this run can no longer apply to anything.
            while (nextRange < ranges.Count && ranges[nextRange].Start + ranges[nextRange].Length <= runStart)
            {
                nextRange++;
            }

            // A tab or break carries no splittable text; style it whole when a range covers it.
            if (run.IsTab || run.IsBreak)
            {
                if (nextRange < ranges.Count && ranges[nextRange].Start < runEnd)
                {
                    run.SetFont(fontName);
                }
                newRuns.Add(run);
                continue;
            }

            // Walk the ranges overlapping this run, emitting unstyled and styled slices in order.
            var cursor = runStart;
            var styledAnything = false;

            for (var i = nextRange; i < ranges.Count && ranges[i].Start < runEnd; i++)
            {
                var (rangeStart, rangeLength) = ranges[i];
                var overlapStart = Math.Max(rangeStart, cursor);
                var overlapEnd = Math.Min(rangeStart + rangeLength, runEnd);
                if (overlapEnd <= overlapStart) continue;

                if (overlapStart > cursor)
                {
                    newRuns.Add(run.CloneWithText(runText[(cursor - runStart)..(overlapStart - runStart)]));
                }

                // A whole-run match keeps the run itself, so nothing about it can be lost.
                if (overlapStart == runStart && overlapEnd == runEnd)
                {
                    run.SetFont(fontName);
                    newRuns.Add(run);
                }
                else
                {
                    var styled = run.CloneWithText(runText[(overlapStart - runStart)..(overlapEnd - runStart)]);
                    styled.SetFont(fontName);
                    newRuns.Add(styled);
                }

                cursor = overlapEnd;
                styledAnything = true;
            }

            if (!styledAnything)
            {
                // Untouched: keep the original run object rather than a copy.
                newRuns.Add(run);
                continue;
            }

            if (cursor < runEnd)
            {
                newRuns.Add(run.CloneWithText(runText[(cursor - runStart)..]));
            }
        }

        node.Runs.Clear();
        node.Runs.AddRange(newRuns);
        node.MarkRunsChanged();
    }

    #endregion
}
