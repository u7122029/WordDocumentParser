namespace WordDocumentParser;

/// <summary>
/// Represents text formatting properties for a run of text
/// </summary>
public class RunFormatting
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public string? UnderlineStyle { get; set; } // Single, Double, Wave, etc.
    public bool Strike { get; set; }
    public bool DoubleStrike { get; set; }
    public string? FontFamily { get; set; }
    public string? FontFamilyAscii { get; set; }
    public string? FontFamilyEastAsia { get; set; }
    public string? FontFamilyComplexScript { get; set; }
    public string? FontSize { get; set; } // In half-points (e.g., "24" = 12pt)
    public string? FontSizeComplexScript { get; set; }
    public string? Color { get; set; } // Hex color without #
    public string? Highlight { get; set; } // Highlight color name
    public bool Superscript { get; set; }
    public bool Subscript { get; set; }
    public bool SmallCaps { get; set; }
    public bool AllCaps { get; set; }
    public string? Shading { get; set; } // Background shading
    public string? StyleId { get; set; } // Character style reference

    public bool HasFormatting =>
        Bold || Italic || Underline || Strike || DoubleStrike ||
        FontFamily is not null || FontSize is not null || Color is not null ||
        Highlight is not null || Superscript || Subscript || SmallCaps || AllCaps;

    public RunFormatting Clone() => new()
    {
        Bold = Bold,
        Italic = Italic,
        Underline = Underline,
        UnderlineStyle = UnderlineStyle,
        Strike = Strike,
        DoubleStrike = DoubleStrike,
        FontFamily = FontFamily,
        FontFamilyAscii = FontFamilyAscii,
        FontFamilyEastAsia = FontFamilyEastAsia,
        FontFamilyComplexScript = FontFamilyComplexScript,
        FontSize = FontSize,
        FontSizeComplexScript = FontSizeComplexScript,
        Color = Color,
        Highlight = Highlight,
        Superscript = Superscript,
        Subscript = Subscript,
        SmallCaps = SmallCaps,
        AllCaps = AllCaps,
        Shading = Shading,
        StyleId = StyleId
    };
}

/// <summary>
/// Represents a run of text with its formatting
/// </summary>
public class FormattedRun
{
    public string Text { get; set; } = string.Empty;
    public RunFormatting Formatting { get; set; } = new();
    public bool IsTab { get; set; }
    public bool IsBreak { get; set; }
    public string? BreakType { get; set; } // TextWrapping, Page, Column

    public FormattedRun() { }
    public FormattedRun(string text) => Text = text;
    public FormattedRun(string text, RunFormatting formatting) => (Text, Formatting) = (text, formatting);
}

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

/// <summary>
/// Represents border formatting
/// </summary>
public class BorderFormatting
{
    public string? Style { get; set; } // Single, Double, Dashed, etc.
    public string? Size { get; set; } // In eighths of a point
    public string? Color { get; set; }
    public string? Space { get; set; } // Space between border and content

    public BorderFormatting Clone() => new()
    {
        Style = Style,
        Size = Size,
        Color = Color,
        Space = Space
    };
}

/// <summary>
/// Represents table formatting properties
/// </summary>
public class TableFormatting
{
    public string? Width { get; set; }
    public string? WidthType { get; set; } // Pct, Dxa, Auto
    public string? Alignment { get; set; } // Left, Center, Right
    public string? IndentFromLeft { get; set; }
    public BorderFormatting? TopBorder { get; set; }
    public BorderFormatting? BottomBorder { get; set; }
    public BorderFormatting? LeftBorder { get; set; }
    public BorderFormatting? RightBorder { get; set; }
    public BorderFormatting? InsideHorizontalBorder { get; set; }
    public BorderFormatting? InsideVerticalBorder { get; set; }
    public string? CellMarginTop { get; set; }
    public string? CellMarginBottom { get; set; }
    public string? CellMarginLeft { get; set; }
    public string? CellMarginRight { get; set; }
    public List<string>? GridColumnWidths { get; set; }

    public TableFormatting Clone() => new()
    {
        Width = Width,
        WidthType = WidthType,
        Alignment = Alignment,
        IndentFromLeft = IndentFromLeft,
        TopBorder = TopBorder?.Clone(),
        BottomBorder = BottomBorder?.Clone(),
        LeftBorder = LeftBorder?.Clone(),
        RightBorder = RightBorder?.Clone(),
        InsideHorizontalBorder = InsideHorizontalBorder?.Clone(),
        InsideVerticalBorder = InsideVerticalBorder?.Clone(),
        CellMarginTop = CellMarginTop,
        CellMarginBottom = CellMarginBottom,
        CellMarginLeft = CellMarginLeft,
        CellMarginRight = CellMarginRight,
        GridColumnWidths = GridColumnWidths is not null ? [.. GridColumnWidths] : null
    };
}

/// <summary>
/// Represents table row formatting
/// </summary>
public class TableRowFormatting
{
    public string? Height { get; set; }
    public string? HeightRule { get; set; } // Auto, Exact, AtLeast
    public bool IsHeader { get; set; }
    public bool CantSplit { get; set; }

    public TableRowFormatting Clone() => new()
    {
        Height = Height,
        HeightRule = HeightRule,
        IsHeader = IsHeader,
        CantSplit = CantSplit
    };
}

/// <summary>
/// Represents table cell formatting
/// </summary>
public class TableCellFormatting
{
    public string? Width { get; set; }
    public string? WidthType { get; set; }
    public int GridSpan { get; set; } = 1;
    public string? VerticalMerge { get; set; } // Restart, Continue, null
    public string? VerticalAlignment { get; set; } // Top, Center, Bottom
    public string? ShadingFill { get; set; }
    public string? ShadingColor { get; set; }
    public string? ShadingPattern { get; set; }
    public BorderFormatting? TopBorder { get; set; }
    public BorderFormatting? BottomBorder { get; set; }
    public BorderFormatting? LeftBorder { get; set; }
    public BorderFormatting? RightBorder { get; set; }
    public string? TextDirection { get; set; }
    public bool NoWrap { get; set; }

    public TableCellFormatting Clone() => new()
    {
        Width = Width,
        WidthType = WidthType,
        GridSpan = GridSpan,
        VerticalMerge = VerticalMerge,
        VerticalAlignment = VerticalAlignment,
        ShadingFill = ShadingFill,
        ShadingColor = ShadingColor,
        ShadingPattern = ShadingPattern,
        TopBorder = TopBorder?.Clone(),
        BottomBorder = BottomBorder?.Clone(),
        LeftBorder = LeftBorder?.Clone(),
        RightBorder = RightBorder?.Clone(),
        TextDirection = TextDirection,
        NoWrap = NoWrap
    };
}

/// <summary>
/// Represents a hyperlink with its URL and formatting
/// </summary>
public class HyperlinkData
{
    public string Text { get; set; } = string.Empty;
    public string? RelationshipId { get; set; }
    public string? Url { get; set; }
    public string? Anchor { get; set; } // For internal document links
    public string? Tooltip { get; set; }
    public List<FormattedRun> Runs { get; set; } = [];
}

/// <summary>
/// Extended image data with positioning information
/// </summary>
public class ImageFormatting
{
    public bool IsInline { get; set; } = true;
    public string? WrapType { get; set; } // None, Square, Tight, Through, TopAndBottom
    public long? DistanceFromTop { get; set; }
    public long? DistanceFromBottom { get; set; }
    public long? DistanceFromLeft { get; set; }
    public long? DistanceFromRight { get; set; }
    public string? HorizontalPosition { get; set; }
    public string? VerticalPosition { get; set; }
    public string? HorizontalRelativeTo { get; set; }
    public string? VerticalRelativeTo { get; set; }
    public long? OffsetX { get; set; }
    public long? OffsetY { get; set; }
    public bool AllowOverlap { get; set; }
    public bool BehindDocument { get; set; }
    public bool LayoutInCell { get; set; }
    public bool Locked { get; set; }
    public long? RelativeHeight { get; set; }
}
