namespace WordDocumentParser.Models.Formatting;

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