using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Tables;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for changing fonts on runs, text spans, and paragraphs.
/// Note: These methods change the actual font family applied to text, not the paragraph style ID.
/// </summary>
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
        run.Formatting ??= new RunFormatting();
        run.Formatting.FontFamily = fontName;
        run.Formatting.FontFamilyAscii = fontName;
        // Also set for other character sets for consistency
        run.Formatting.FontFamilyEastAsia = fontName;
        run.Formatting.FontFamilyComplexScript = fontName;
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
        run.Formatting ??= new RunFormatting();
        run.Formatting.FontFamilyAscii = ascii;
        run.Formatting.FontFamily = highAnsi ?? ascii;
        run.Formatting.FontFamilyEastAsia = eastAsia;
        run.Formatting.FontFamilyComplexScript = complexScript;
    }

    /// <summary>
    /// Gets the font family name from a formatted run.
    /// </summary>
    /// <param name="run">The run to check</param>
    /// <returns>The font family name, or null if not set</returns>
    public static string? GetFont(this FormattedRun run)
        => run.Formatting?.FontFamilyAscii ?? run.Formatting?.FontFamily;

    #endregion

    #region Paragraph-level font changes

    /// <summary>
    /// Sets the font family for all runs in a paragraph node.
    /// This changes the actual font applied to all text, not the paragraph style.
    /// </summary>
    /// <param name="node">The paragraph node to modify</param>
    /// <param name="fontName">The font family name (e.g., "Calibri", "Arial")</param>
    public static void SetParagraphFont(this DocumentNode node, string fontName)
    {
        if (node.Type is not (ContentType.Paragraph or ContentType.Heading or ContentType.ListItem))
            return;

        // Ensure we have the text content
        var text = node.GetText();
        if (string.IsNullOrEmpty(text) && string.IsNullOrEmpty(node.Text))
            return;

        // If the node has formatted runs, update each one
        if (node.HasFormattedRuns)
        {
            foreach (var run in node.Runs)
            {
                run.SetFont(fontName);
            }
        }
        else
        {
            // Create a formatted run from the plain text with the font applied
            var textToUse = !string.IsNullOrEmpty(node.Text) ? node.Text : text;
            var run = new FormattedRun(textToUse);
            run.SetFont(fontName);
            node.Runs.Add(run);
        }

        // Clear OriginalXml so the writer generates clean XML from formatted runs
        // This is safer than trying to modify XML with regex which can create malformed XML
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            node.OriginalXml = null;
        }
    }

    /// <summary>
    /// Sets the font family for all paragraphs in a document.
    /// </summary>
    /// <param name="document">The document to modify</param>
    /// <param name="fontName">The font family name</param>
    /// <returns>The number of paragraphs modified</returns>
    public static int SetDocumentFont(this WordDocument document, string fontName)
        => document.Root.SetDocumentFont(fontName);

    /// <summary>
    /// Sets the font family for all paragraphs under a node.
    /// </summary>
    /// <param name="root">The root node to start from</param>
    /// <param name="fontName">The font family name</param>
    /// <returns>The number of paragraphs modified</returns>
    public static int SetDocumentFont(this DocumentNode root, string fontName)
    {
        var count = 0;
        foreach (var node in root.FindAll(n => n.Type is ContentType.Paragraph or ContentType.Heading or ContentType.ListItem))
        {
            node.SetParagraphFont(fontName);
            count++;
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
    public static int SetFontForText(this DocumentNode node, string searchText, string fontName, bool allOccurrences = false)
    {
        if (node.Type is not (ContentType.Paragraph or ContentType.Heading or ContentType.ListItem))
            return 0;

        if (string.IsNullOrEmpty(searchText))
            return 0;

        // Ensure we have formatted runs to work with
        if (!node.HasFormattedRuns && !string.IsNullOrEmpty(node.Text))
        {
            node.Runs.Add(new FormattedRun(node.Text));
        }

        var count = 0;
        var modified = true;

        while (modified)
        {
            modified = false;
            var fullText = string.Concat(node.Runs.Select(r => r.Text));
            var index = fullText.IndexOf(searchText, StringComparison.Ordinal);

            if (index >= 0)
            {
                ApplyFontToRange(node, index, searchText.Length, fontName);
                count++;
                modified = allOccurrences;
            }
        }

        // Update OriginalXml - for span changes, we need to rebuild or mark as modified
        if (count > 0 && !string.IsNullOrEmpty(node.OriginalXml))
        {
            // Clear OriginalXml so the writer uses the formatted runs instead
            node.OriginalXml = null;
        }

        return count;
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

        // Ensure we have formatted runs to work with
        if (!node.HasFormattedRuns && !string.IsNullOrEmpty(node.Text))
        {
            node.Runs.Add(new FormattedRun(node.Text));
        }

        var fullText = string.Concat(node.Runs.Select(r => r.Text));
        if (startIndex + length > fullText.Length)
            return false;

        ApplyFontToRange(node, startIndex, length, fontName);

        // Clear OriginalXml so the writer uses the formatted runs
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            node.OriginalXml = null;
        }

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

        if (count > 0 && !string.IsNullOrEmpty(node.OriginalXml))
        {
            node.OriginalXml = null;
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

        if (node.HasFormattedRuns)
        {
            foreach (var run in node.Runs)
            {
                var font = run.GetFont();
                if (!string.IsNullOrEmpty(font))
                {
                    fonts.Add(font);
                }
            }
        }

        return fonts;
    }

    /// <summary>
    /// Gets all unique font families used in a document.
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

        foreach (var node in root.FindAll(_ => true))
        {
            // Get fonts from the node itself
            foreach (var font in node.GetFontsUsed())
            {
                fonts.Add(font);
            }

            // If this is a table, also check cell content
            if (node.Type == Core.ContentType.Table)
            {
                var tableData = node.GetTableData();
                if (tableData != null)
                {
                    foreach (var row in tableData.Rows)
                    {
                        foreach (var cell in row.Cells)
                        {
                            foreach (var content in cell.Content)
                            {
                                foreach (var font in content.GetFontsUsed())
                                {
                                    fonts.Add(font);
                                }
                            }
                        }
                    }
                }
            }
        }

        return fonts;
    }

    /// <summary>
    /// Replaces one font with another throughout a document.
    /// </summary>
    /// <param name="document">The document to modify</param>
    /// <param name="fromFont">The font to replace</param>
    /// <param name="toFont">The font to replace with</param>
    /// <returns>The number of runs modified</returns>
    public static int ReplaceFont(this WordDocument document, string fromFont, string toFont)
        => document.Root.ReplaceFont(fromFont, toFont);

    /// <summary>
    /// Replaces one font with another under a node.
    /// </summary>
    /// <param name="root">The root node to start from</param>
    /// <param name="fromFont">The font to replace</param>
    /// <param name="toFont">The font to replace with</param>
    /// <returns>The number of runs modified</returns>
    public static int ReplaceFont(this DocumentNode root, string fromFont, string toFont)
    {
        var count = 0;

        foreach (var node in root.FindAll(n => n.HasFormattedRuns))
        {
            var nodeModified = false;
            foreach (var run in node.Runs)
            {
                var currentFont = run.GetFont();
                if (currentFont != null && currentFont.Equals(fromFont, StringComparison.OrdinalIgnoreCase))
                {
                    run.SetFont(toFont);
                    count++;
                    nodeModified = true;
                }
            }

            // Clear OriginalXml if we made changes so writer generates clean XML
            if (nodeModified && !string.IsNullOrEmpty(node.OriginalXml))
            {
                node.OriginalXml = null;
            }
        }

        return count;
    }

    #endregion

    #region Private helpers

    /// <summary>
    /// Applies a font to a character range by splitting runs as needed.
    /// </summary>
    private static void ApplyFontToRange(DocumentNode node, int startIndex, int length, string fontName)
    {
        var newRuns = new List<FormattedRun>();
        var currentPos = 0;
        var endIndex = startIndex + length;

        foreach (var run in node.Runs)
        {
            var runStart = currentPos;
            var runEnd = currentPos + run.Text.Length;

            if (runEnd <= startIndex || runStart >= endIndex)
            {
                // Run is entirely outside the target range - keep as is
                newRuns.Add(run);
            }
            else if (runStart >= startIndex && runEnd <= endIndex)
            {
                // Run is entirely inside the target range - apply font
                run.SetFont(fontName);
                newRuns.Add(run);
            }
            else
            {
                // Run partially overlaps - need to split
                var overlapStart = Math.Max(startIndex, runStart);
                var overlapEnd = Math.Min(endIndex, runEnd);

                // Part before the overlap
                if (runStart < overlapStart)
                {
                    var beforeText = run.Text[..(overlapStart - runStart)];
                    var beforeRun = new FormattedRun(beforeText, run.Formatting.Clone());
                    newRuns.Add(beforeRun);
                }

                // The overlapping part with new font
                var overlapText = run.Text[(overlapStart - runStart)..(overlapEnd - runStart)];
                var overlapRun = new FormattedRun(overlapText, run.Formatting.Clone());
                overlapRun.SetFont(fontName);
                newRuns.Add(overlapRun);

                // Part after the overlap
                if (runEnd > overlapEnd)
                {
                    var afterText = run.Text[(overlapEnd - runStart)..];
                    var afterRun = new FormattedRun(afterText, run.Formatting.Clone());
                    newRuns.Add(afterRun);
                }
            }

            currentPos = runEnd;
        }

        node.Runs.Clear();
        node.Runs.AddRange(newRuns);
    }

    #endregion
}
