namespace WordDocumentParser.FormattingModels;

/// <summary>
/// Represents an item in a dropdown list or combo box content control
/// </summary>
public class ContentControlListItem
{
    /// <summary>
    /// Display text for the item
    /// </summary>
    public string? DisplayText { get; set; }

    /// <summary>
    /// Value of the item
    /// </summary>
    public string? Value { get; set; }

    public ContentControlListItem Clone() => new()
    {
        DisplayText = DisplayText,
        Value = Value
    };
}