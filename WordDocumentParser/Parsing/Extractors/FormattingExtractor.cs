using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Parsing.Extractors;

/// <summary>
/// Extracts formatting information from OpenXML elements.
/// </summary>
internal static class FormattingExtractor
{
    /// <summary>
    /// Extracts run-level formatting from RunProperties.
    /// </summary>
    public static RunFormatting ExtractRunFormatting(RunProperties? rPr)
    {
        var formatting = new RunFormatting();
        if (rPr is null) return formatting;

        // Bold
        formatting.Bold = rPr.Bold is not null && (rPr.Bold.Val is null || rPr.Bold.Val.Value);

        // Italic
        formatting.Italic = rPr.Italic is not null && (rPr.Italic.Val is null || rPr.Italic.Val.Value);

        // Underline
        if (rPr.Underline is not null)
        {
            formatting.Underline = rPr.Underline.Val?.Value != UnderlineValues.None;
            formatting.UnderlineStyle = rPr.Underline.Val?.Value.ToString();
        }

        // Strike
        formatting.Strike = rPr.Strike is not null && (rPr.Strike.Val is null || rPr.Strike.Val.Value);
        formatting.DoubleStrike = rPr.DoubleStrike is not null && (rPr.DoubleStrike.Val is null || rPr.DoubleStrike.Val.Value);

        // Font
        var fonts = rPr.RunFonts;
        if (fonts is not null)
        {
            formatting.FontFamily = fonts.HighAnsi?.Value;
            formatting.FontFamilyAscii = fonts.Ascii?.Value;
            formatting.FontFamilyEastAsia = fonts.EastAsia?.Value;
            formatting.FontFamilyComplexScript = fonts.ComplexScript?.Value;
        }

        // Font size
        formatting.FontSize = rPr.FontSize?.Val?.Value;
        formatting.FontSizeComplexScript = rPr.FontSizeComplexScript?.Val?.Value;

        // Color
        formatting.Color = rPr.Color?.Val?.Value;

        // Highlight
        formatting.Highlight = rPr.Highlight?.Val?.Value.ToString();

        // Superscript/Subscript
        var vertAlign = rPr.VerticalTextAlignment?.Val?.Value;
        if (vertAlign.HasValue)
        {
            formatting.Superscript = vertAlign.Value == VerticalPositionValues.Superscript;
            formatting.Subscript = vertAlign.Value == VerticalPositionValues.Subscript;
        }

        // Caps
        formatting.SmallCaps = rPr.SmallCaps is not null && (rPr.SmallCaps.Val is null || rPr.SmallCaps.Val.Value);
        formatting.AllCaps = rPr.Caps is not null && (rPr.Caps.Val is null || rPr.Caps.Val.Value);

        // Shading
        formatting.Shading = rPr.Shading?.Fill?.Value;

        // Character style
        formatting.StyleId = rPr.RunStyle?.Val?.Value;

        return formatting;
    }

    /// <summary>
    /// Extracts paragraph-level formatting from a Paragraph element.
    /// </summary>
    public static ParagraphFormatting ExtractParagraphFormatting(Paragraph para, ParsingContext context)
    {
        var formatting = new ParagraphFormatting();
        var pPr = para.ParagraphProperties;
        if (pPr is null) return formatting;

        // Style
        formatting.StyleId = pPr.ParagraphStyleId?.Val?.Value;

        // Alignment
        formatting.Alignment = pPr.Justification?.Val?.Value.ToString();

        // Indentation
        var ind = pPr.Indentation;
        if (ind is not null)
        {
            formatting.IndentLeft = ind.Left?.Value;
            formatting.IndentRight = ind.Right?.Value;
            formatting.IndentFirstLine = ind.FirstLine?.Value;
            formatting.IndentHanging = ind.Hanging?.Value;
        }

        // Spacing
        var spacing = pPr.SpacingBetweenLines;
        if (spacing is not null)
        {
            formatting.SpacingBefore = spacing.Before?.Value;
            formatting.SpacingAfter = spacing.After?.Value;
            formatting.LineSpacing = spacing.Line?.Value;
            formatting.LineSpacingRule = spacing.LineRule?.Value.ToString();
        }

        // Keep with next/keep lines
        formatting.KeepNext = pPr.KeepNext is not null;
        formatting.KeepLines = pPr.KeepLines is not null;

        // Page break before
        formatting.PageBreakBefore = pPr.PageBreakBefore is not null;

        // Widow control
        formatting.WidowControl = pPr.WidowControl is not null;

        // Outline level
        formatting.OutlineLevel = pPr.OutlineLevel?.Val?.Value.ToString();

        // Shading
        var shading = pPr.Shading;
        if (shading is not null)
        {
            formatting.ShadingFill = shading.Fill?.Value;
            formatting.ShadingColor = shading.Color?.Value;
        }

        // Borders
        var borders = pPr.ParagraphBorders;
        if (borders is not null)
        {
            formatting.TopBorder = ExtractBorderFormatting(borders.TopBorder);
            formatting.BottomBorder = ExtractBorderFormatting(borders.BottomBorder);
            formatting.LeftBorder = ExtractBorderFormatting(borders.LeftBorder);
            formatting.RightBorder = ExtractBorderFormatting(borders.RightBorder);
        }

        // Numbering
        var numPr = pPr.NumberingProperties;
        if (numPr is not null)
        {
            formatting.NumberingId = numPr.NumberingId?.Val?.Value;
            formatting.NumberingLevel = numPr.NumberingLevelReference?.Val?.Value;
        }

        return formatting;
    }

    /// <summary>
    /// Extracts border formatting from a border element.
    /// </summary>
    public static BorderFormatting? ExtractBorderFormatting(BorderType? border)
    {
        if (border is null) return null;

        return new BorderFormatting
        {
            Style = border.Val?.Value.ToString(),
            Size = border.Size?.Value.ToString(),
            Color = border.Color?.Value,
            Space = border.Space?.Value.ToString()
        };
    }
}
