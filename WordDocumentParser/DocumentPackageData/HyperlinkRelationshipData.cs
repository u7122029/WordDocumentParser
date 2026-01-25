namespace WordDocumentParser.DocumentPackageData;

/// <summary>
/// Stores hyperlink relationship data
/// </summary>
public class HyperlinkRelationshipData
{
    public string Url { get; set; } = string.Empty;
    public bool IsExternal { get; set; } = true;
}