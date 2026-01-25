namespace WordDocumentParser.DocumentPackageData;

/// <summary>
/// Stores image part data
/// </summary>
public class ImagePartData
{
    public string ContentType { get; set; } = string.Empty;
    public byte[] Data { get; set; } = [];
    public string OriginalRelationshipId { get; set; } = string.Empty;
    public string? OriginalUri { get; set; }
}
