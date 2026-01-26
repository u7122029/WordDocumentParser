namespace WordDocumentParser.Models.ContentControls;

/// <summary>
/// Represents a document property field in the document
/// </summary>
public class DocumentPropertyField
{
    /// <summary>
    /// Type of property field (Core, Extended, Custom)
    /// </summary>
    public DocumentPropertyType PropertyType { get; set; }

    /// <summary>
    /// Name of the property
    /// </summary>
    public string PropertyName { get; set; } = string.Empty;

    /// <summary>
    /// Current value of the property
    /// </summary>
    public string? Value { get; set; }

    /// <summary>
    /// The original field code
    /// </summary>
    public string? FieldCode { get; set; }

    /// <summary>
    /// Gets a string representation showing metadata about this property field
    /// </summary>
    public string ToMetadataString()
        => $"[DocProperty:{PropertyType}/{PropertyName}=\"{Value ?? "(empty)"}\"]";
}