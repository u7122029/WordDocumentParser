using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Models.Images;

/// <summary>
/// Represents image data extracted from a Word document.
/// </summary>
public class ImageData
{
    /// <summary>Relationship ID of the image in the document package</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Name/title of the image</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>MIME content type (e.g., "image/png", "image/jpeg")</summary>
    public string ContentType { get; set; } = string.Empty;

    /// <summary>Raw image binary data</summary>
    public byte[]? Data { get; set; }

    /// <summary>Display width in inches</summary>
    public double WidthInches { get; set; }

    /// <summary>Display height in inches</summary>
    public double HeightInches { get; set; }

    /// <summary>Alt text for accessibility</summary>
    public string? AltText { get; set; }

    /// <summary>Extended description</summary>
    public string? Description { get; set; }

    /// <summary>Width in EMUs (English Metric Units) for precise round-trip</summary>
    public long WidthEmu { get; set; }

    /// <summary>Height in EMUs (English Metric Units) for precise round-trip</summary>
    public long HeightEmu { get; set; }

    /// <summary>Image positioning and formatting for round-trip fidelity</summary>
    public ImageFormatting? Formatting { get; set; }
}
