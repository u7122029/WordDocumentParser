namespace WordDocumentParser.FormattingModels;

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