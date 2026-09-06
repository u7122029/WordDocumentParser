namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// A hyperlink found in a paragraph, with its target and display runs.
/// </summary>
public class HyperlinkData
{
    /// <summary>The link's display text.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Relationship ID of the link target, for external links.</summary>
    public string? RelationshipId { get; set; }

    /// <summary>Resolved target URL, for external links.</summary>
    public string? Url { get; set; }

    /// <summary>Bookmark name, for links within the document.</summary>
    public string? Anchor { get; set; }

    /// <summary>Tooltip shown on hover.</summary>
    public string? Tooltip { get; set; }

    /// <summary>The formatted runs that make up the link's display text.</summary>
    public List<FormattedRun> Runs { get; set; } = [];
}
