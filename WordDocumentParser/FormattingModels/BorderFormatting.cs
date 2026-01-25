namespace WordDocumentParser.FormattingModels;

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