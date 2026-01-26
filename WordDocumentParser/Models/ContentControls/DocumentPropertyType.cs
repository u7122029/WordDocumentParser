namespace WordDocumentParser.Models.ContentControls;

/// <summary>
/// Types of document properties
/// </summary>
public enum DocumentPropertyType
{
    Core,       // Title, Subject, Author, etc.
    Extended,   // Company, Manager, etc.
    Custom      // User-defined properties
}
