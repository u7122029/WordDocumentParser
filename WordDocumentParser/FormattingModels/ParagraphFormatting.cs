namespace WordDocumentParser.FormattingModels;

/// <summary>
/// Represents paragraph formatting properties
/// </summary>
public class ParagraphFormatting
{
    public string? StyleId { get; set; }
    public string? Alignment { get; set; } // Left, Center, Right, Both (justify)
    public string? IndentLeft { get; set; } // In twips
    public string? IndentRight { get; set; }
    public string? IndentFirstLine { get; set; }
    public string? IndentHanging { get; set; }
    public string? SpacingBefore { get; set; } // In twips
    public string? SpacingAfter { get; set; }
    public string? LineSpacing { get; set; }
    public string? LineSpacingRule { get; set; } // Auto, Exact, AtLeast
    public bool KeepNext { get; set; }
    public bool KeepLines { get; set; }
    public bool PageBreakBefore { get; set; }
    public bool WidowControl { get; set; }
    public string? OutlineLevel { get; set; }
    public string? ShadingFill { get; set; }
    public string? ShadingColor { get; set; }
    public BorderFormatting? TopBorder { get; set; }
    public BorderFormatting? BottomBorder { get; set; }
    public BorderFormatting? LeftBorder { get; set; }
    public BorderFormatting? RightBorder { get; set; }

    // Numbering properties
    public int? NumberingId { get; set; }
    public int? NumberingLevel { get; set; }

    public bool HasFormatting =>
        StyleId is not null || Alignment is not null ||
        IndentLeft is not null || IndentRight is not null ||
        IndentFirstLine is not null || IndentHanging is not null ||
        SpacingBefore is not null || SpacingAfter is not null ||
        LineSpacing is not null || KeepNext || KeepLines ||
        PageBreakBefore || ShadingFill is not null;

    public ParagraphFormatting Clone() => new()
    {
        StyleId = StyleId,
        Alignment = Alignment,
        IndentLeft = IndentLeft,
        IndentRight = IndentRight,
        IndentFirstLine = IndentFirstLine,
        IndentHanging = IndentHanging,
        SpacingBefore = SpacingBefore,
        SpacingAfter = SpacingAfter,
        LineSpacing = LineSpacing,
        LineSpacingRule = LineSpacingRule,
        KeepNext = KeepNext,
        KeepLines = KeepLines,
        PageBreakBefore = PageBreakBefore,
        WidowControl = WidowControl,
        OutlineLevel = OutlineLevel,
        ShadingFill = ShadingFill,
        ShadingColor = ShadingColor,
        TopBorder = TopBorder?.Clone(),
        BottomBorder = BottomBorder?.Clone(),
        LeftBorder = LeftBorder?.Clone(),
        RightBorder = RightBorder?.Clone(),
        NumberingId = NumberingId,
        NumberingLevel = NumberingLevel
    };
}