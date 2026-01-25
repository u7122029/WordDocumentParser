namespace WordDocumentParser.FormattingModels;

/// <summary>
/// Represents table row formatting
/// </summary>
public class TableRowFormatting
{
    public string? Height { get; set; }
    public string? HeightRule { get; set; } // Auto, Exact, AtLeast
    public bool IsHeader { get; set; }
    public bool CantSplit { get; set; }

    public TableRowFormatting Clone() => new()
    {
        Height = Height,
        HeightRule = HeightRule,
        IsHeader = IsHeader,
        CantSplit = CantSplit
    };
}