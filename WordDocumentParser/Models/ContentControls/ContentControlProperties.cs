using WordDocumentParser.Core;

namespace WordDocumentParser.Models.ContentControls;

/// <summary>
/// Properties of a structured document tag (content control).
/// </summary>
/// <remarks>
/// Assignments are tracked (see <see cref="TrackedModel"/>). The writer rewrites the control's
/// properties in the original SDT only when a caller has actually changed one; an untouched control
/// passes through with its definition intact, including extension-namespace children such as the
/// <c>w14:checkbox</c> element that carries a checkbox's state.
/// </remarks>
public class ContentControlProperties : TrackedModel
{
    private int? _id;
    private string? _tag;
    private string? _alias;
    private ContentControlType _type = ContentControlType.Unknown;
    private string? _placeholderText;
    private bool _lockContentControl;
    private bool _lockContents;
    private string? _dataBindingPrefixMappings;
    private string? _dataBindingXPath;
    private string? _dataBindingStoreItemId;
    private string? _dateFormat;
    private string? _dateLocale;
    private DateTime? _dateValue;
    private List<ContentControlListItem> _listItems = [];
    private bool _showingPlaceholder;
    private string? _value;
    private bool? _isChecked;
    private string? _color;
    private string? _appearance;

    /// <summary>Unique identifier for the content control.</summary>
    public int? Id { get => _id; set => Set(ref _id, value); }

    /// <summary>Tag used for programmatic identification.</summary>
    public string? Tag { get => _tag; set => Set(ref _tag, value); }

    /// <summary>Alias/title displayed in the Word UI.</summary>
    public string? Alias { get => _alias; set => Set(ref _alias, value); }

    /// <summary>The kind of content control.</summary>
    public ContentControlType Type { get => _type; set => Set(ref _type, value); }

    /// <summary>Placeholder text shown when the control is empty.</summary>
    public string? PlaceholderText { get => _placeholderText; set => Set(ref _placeholderText, value); }

    /// <summary>Whether the control itself is protected from deletion.</summary>
    public bool LockContentControl { get => _lockContentControl; set => Set(ref _lockContentControl, value); }

    /// <summary>Whether the control's contents are protected from editing.</summary>
    public bool LockContents { get => _lockContents; set => Set(ref _lockContents, value); }

    /// <summary>Namespace prefix mappings for the data binding.</summary>
    public string? DataBindingPrefixMappings { get => _dataBindingPrefixMappings; set => Set(ref _dataBindingPrefixMappings, value); }

    /// <summary>XPath for the data binding.</summary>
    public string? DataBindingXPath { get => _dataBindingXPath; set => Set(ref _dataBindingXPath, value); }

    /// <summary>Store item ID for the data binding.</summary>
    public string? DataBindingStoreItemId { get => _dataBindingStoreItemId; set => Set(ref _dataBindingStoreItemId, value); }

    /// <summary>Date format string, for date controls.</summary>
    public string? DateFormat { get => _dateFormat; set => Set(ref _dateFormat, value); }

    /// <summary>Locale identifier, for date controls.</summary>
    public string? DateLocale { get => _dateLocale; set => Set(ref _dateLocale, value); }

    /// <summary>Current value, for date controls.</summary>
    public DateTime? DateValue { get => _dateValue; set => Set(ref _dateValue, value); }

    /// <summary>Items offered by a dropdown or combo box control.</summary>
    public List<ContentControlListItem> ListItems { get => _listItems; set => Set(ref _listItems, value ?? []); }

    /// <summary>Whether the control is currently showing its placeholder.</summary>
    public bool ShowingPlaceholder { get => _showingPlaceholder; set => Set(ref _showingPlaceholder, value); }

    /// <summary>The control's current text value.</summary>
    public string? Value { get => _value; set => Set(ref _value, value); }

    /// <summary>Checked state, for checkbox controls.</summary>
    public bool? IsChecked { get => _isChecked; set => Set(ref _isChecked, value); }

    /// <summary>Border colour of the control.</summary>
    public string? Color { get => _color; set => Set(ref _color, value); }

    /// <summary>Appearance setting, for example <c>"boundingBox"</c>, <c>"tags"</c>, <c>"hidden"</c>.</summary>
    public string? Appearance { get => _appearance; set => Set(ref _appearance, value); }

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public ContentControlProperties Clone()
    {
        var clone = new ContentControlProperties
        {
            _id = _id,
            _tag = _tag,
            _alias = _alias,
            _type = _type,
            _placeholderText = _placeholderText,
            _lockContentControl = _lockContentControl,
            _lockContents = _lockContents,
            _dataBindingPrefixMappings = _dataBindingPrefixMappings,
            _dataBindingXPath = _dataBindingXPath,
            _dataBindingStoreItemId = _dataBindingStoreItemId,
            _dateFormat = _dateFormat,
            _dateLocale = _dateLocale,
            _dateValue = _dateValue,
            _listItems = [.. _listItems.Select(i => i.Clone())],
            _showingPlaceholder = _showingPlaceholder,
            _value = _value,
            _isChecked = _isChecked,
            _color = _color,
            _appearance = _appearance
        };
        clone.CopyChangesFrom(this);
        return clone;
    }

    /// <summary>
    /// Gets a string representation showing metadata about this content control.
    /// </summary>
    /// <returns>A single-line description.</returns>
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
