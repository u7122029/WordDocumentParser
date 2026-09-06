namespace WordDocumentParser.Models.Package;

/// <summary>
/// A hyperlink relationship from a package part.
/// </summary>
public class HyperlinkRelationshipData
{
    /// <summary>The relationship target, absolute for external links.</summary>
    public string Url { get; set; } = string.Empty;

    /// <summary>Whether the target is outside the package.</summary>
    public bool IsExternal { get; set; } = true;
}
