namespace WordDocumentParser.FormattingModels;

/// <summary>
/// Represents properties of a structured document tag (content control)
/// </summary>
public class ContentControlProperties
{
    /// <summary>
    /// Unique identifier for the content control
    /// </summary>
    public int? Id { get; set; }

    /// <summary>
    /// Tag for the content control (used for programmatic identification)
    /// </summary>
    public string? Tag { get; set; }

    /// <summary>
    /// Alias/Title displayed in the UI
    /// </summary>
    public string? Alias { get; set; }

    /// <summary>
    /// The type of content control
    /// </summary>
    public ContentControlType Type { get; set; } = ContentControlType.Unknown;

    /// <summary>
    /// Placeholder text shown when the control is empty
    /// </summary>
    public string? PlaceholderText { get; set; }

    /// <summary>
    /// Whether the content control can be deleted
    /// </summary>
    public bool LockContentControl { get; set; }

    /// <summary>
    /// Whether the contents of the content control can be edited
    /// </summary>
    public bool LockContents { get; set; }

    /// <summary>
    /// For document property content controls, the name of the property
    /// </summary>
    public string? DataBindingPrefixMappings { get; set; }

    /// <summary>
    /// XPath for data binding
    /// </summary>
    public string? DataBindingXPath { get; set; }

    /// <summary>
    /// Store item ID for data binding
    /// </summary>
    public string? DataBindingStoreItemId { get; set; }

    /// <summary>
    /// For date content controls, the date format
    /// </summary>
    public string? DateFormat { get; set; }

    /// <summary>
    /// For date content controls, the locale
    /// </summary>
    public string? DateLocale { get; set; }

    /// <summary>
    /// For date content controls, the current date value
    /// </summary>
    public DateTime? DateValue { get; set; }

    /// <summary>
    /// For dropdown/combobox controls, the list of items
    /// </summary>
    public List<ContentControlListItem> ListItems { get; set; } = [];

    /// <summary>
    /// Whether to show the control as a bounding box
    /// </summary>
    public bool ShowingPlaceholder { get; set; }

    /// <summary>
    /// The current value/text of the content control
    /// </summary>
    public string? Value { get; set; }

    /// <summary>
    /// For checkbox controls, whether it's checked
    /// </summary>
    public bool? IsChecked { get; set; }

    /// <summary>
    /// Color of the content control border
    /// </summary>
    public string? Color { get; set; }

    /// <summary>
    /// Appearance setting (BoundingBox, Tags, Hidden, etc.)
    /// </summary>
    public string? Appearance { get; set; }

    public ContentControlProperties Clone() => new()
    {
        Id = Id,
        Tag = Tag,
        Alias = Alias,
        Type = Type,
        PlaceholderText = PlaceholderText,
        LockContentControl = LockContentControl,
        LockContents = LockContents,
        DataBindingPrefixMappings = DataBindingPrefixMappings,
        DataBindingXPath = DataBindingXPath,
        DataBindingStoreItemId = DataBindingStoreItemId,
        DateFormat = DateFormat,
        DateLocale = DateLocale,
        DateValue = DateValue,
        ListItems = [.. ListItems.Select(i => i.Clone())],
        ShowingPlaceholder = ShowingPlaceholder,
        Value = Value,
        IsChecked = IsChecked,
        Color = Color,
        Appearance = Appearance
    };

    /// <summary>
    /// Gets a string representation showing metadata about this content control
    /// </summary>
    public string ToMetadataString()
    {
        var parts = new List<string> { $"[ContentControl:{Type}" };

        if (!string.IsNullOrEmpty(Alias))
            parts.Add($"Alias=\"{Alias}\"");
        else if (!string.IsNullOrEmpty(Tag))
            parts.Add($"Tag=\"{Tag}\"");

        if (!string.IsNullOrEmpty(Value))
            parts.Add($"Value=\"{Value}\"");

        parts[^1] += "]";
        return string.Join(" ", parts);
    }
}