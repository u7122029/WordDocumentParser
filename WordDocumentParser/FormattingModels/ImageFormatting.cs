namespace WordDocumentParser.FormattingModels;

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