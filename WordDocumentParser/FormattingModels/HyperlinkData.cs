namespace WordDocumentParser.FormattingModels;

/// <summary>
/// Represents a hyperlink with its URL and formatting
/// </summary>
public class HyperlinkData
{
    public string Text { get; set; } = string.Empty;
    public string? RelationshipId { get; set; }
    public string? Url { get; set; }
    public string? Anchor { get; set; } // For internal document links
    public string? Tooltip { get; set; }
    public List<FormattedRun> Runs { get; set; } = [];
}