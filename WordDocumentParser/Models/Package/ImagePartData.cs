namespace WordDocumentParser.Models.Package;

/// <summary>
/// A binary media part from the document package.
/// </summary>
/// <remarks>
/// <see cref="Data"/> is shared: every reference to the same image part in a parsed document points
/// at one array, so a picture used thirty times costs one copy rather than thirty. Treat the array
/// as immutable — writing through it changes every reference, and the copies a document clone hands
/// out share it as well.
/// </remarks>
public class ImagePartData
{
    /// <summary>MIME content type of the part, for example <c>"image/png"</c>.</summary>
    public string ContentType { get; set; } = string.Empty;

    /// <summary>
    /// The part's decompressed bytes. Shared between every reference to this part; do not modify
    /// in place.
    /// </summary>
    public byte[] Data { get; set; } = [];

    /// <summary>The relationship ID this part had in the source package.</summary>
    public string OriginalRelationshipId { get; set; } = string.Empty;

    /// <summary>The part URI this part had in the source package.</summary>
    public string? OriginalUri { get; set; }
}
