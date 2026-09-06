namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Positioning and wrapping for an image.
/// </summary>
public class ImageFormatting
{
    /// <summary>True for an inline image, false for one anchored to the page or paragraph.</summary>
    public bool IsInline { get; set; } = true;

    /// <summary>Text wrapping mode: <c>None</c>, <c>Square</c>, <c>Tight</c>, <c>Through</c>, or <c>TopAndBottom</c>.</summary>
    public string? WrapType { get; set; }

    /// <summary>Distance from surrounding text above, in EMUs.</summary>
    public long? DistanceFromTop { get; set; }

    /// <summary>Distance from surrounding text below, in EMUs.</summary>
    public long? DistanceFromBottom { get; set; }

    /// <summary>Distance from surrounding text to the left, in EMUs.</summary>
    public long? DistanceFromLeft { get; set; }

    /// <summary>Distance from surrounding text to the right, in EMUs.</summary>
    public long? DistanceFromRight { get; set; }

    /// <summary>Horizontal offset from the anchor, in EMUs, as text.</summary>
    public string? HorizontalPosition { get; set; }

    /// <summary>Vertical offset from the anchor, in EMUs, as text.</summary>
    public string? VerticalPosition { get; set; }

    /// <summary>What the horizontal position is measured from: <c>Column</c>, <c>Page</c>, or <c>Margin</c>.</summary>
    public string? HorizontalRelativeTo { get; set; }

    /// <summary>What the vertical position is measured from: <c>Paragraph</c>, <c>Page</c>, or <c>Margin</c>.</summary>
    public string? VerticalRelativeTo { get; set; }

    /// <summary>Additional horizontal offset in EMUs.</summary>
    public long? OffsetX { get; set; }

    /// <summary>Additional vertical offset in EMUs.</summary>
    public long? OffsetY { get; set; }

    /// <summary>Whether this image may overlap other floating objects.</summary>
    public bool AllowOverlap { get; set; }

    /// <summary>Whether the image sits behind the text.</summary>
    public bool BehindDocument { get; set; }

    /// <summary>Whether the image is laid out within its containing table cell.</summary>
    public bool LayoutInCell { get; set; }

    /// <summary>Whether the anchor is locked against moving.</summary>
    public bool Locked { get; set; }

    /// <summary>Z-order among floating objects.</summary>
    public long? RelativeHeight { get; set; }
}
