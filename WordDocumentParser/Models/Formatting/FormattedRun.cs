using WordDocumentParser.Core;
using WordDocumentParser.Models.ContentControls;

namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// A run of text with its formatting.
/// </summary>
/// <remarks>
/// Assignments are tracked (see <see cref="TrackedModel"/>) so the writer can tell a run the caller
/// edited from one it merely parsed, and rewrite only the former.
/// </remarks>
public class FormattedRun : TrackedModel
{
    private string _text = string.Empty;
    private RunFormatting _formatting = new();
    private bool _isTab;
    private bool _isBreak;
    private string? _breakType;
    private DocumentPropertyField? _documentPropertyField;
    private ContentControlProperties? _contentControlProperties;

    /// <summary>The run's text.</summary>
    public string Text { get => _text; set => Set(ref _text, value ?? string.Empty); }

    /// <summary>Character formatting applied to this run.</summary>
    public RunFormatting Formatting { get => _formatting; set => Set(ref _formatting, value ?? new RunFormatting()); }

    /// <summary>This run is a tab character rather than text.</summary>
    public bool IsTab { get => _isTab; set => Set(ref _isTab, value); }

    /// <summary>This run is a break rather than text.</summary>
    public bool IsBreak { get => _isBreak; set => Set(ref _isBreak, value); }

    /// <summary>Break kind as an OOXML token: <c>"page"</c>, <c>"column"</c>, <c>"textWrapping"</c>.</summary>
    public string? BreakType { get => _breakType; set => Set(ref _breakType, value); }

    /// <summary>
    /// Field information when this run is the result of a document property field.
    /// </summary>
    public DocumentPropertyField? DocumentPropertyField
    {
        get => _documentPropertyField;
        set => Set(ref _documentPropertyField, value);
    }

    /// <summary>Whether this run is a document property field.</summary>
    public bool IsDocumentPropertyField => DocumentPropertyField is not null;

    /// <summary>
    /// Control properties when this run sits inside an inline content control.
    /// </summary>
    public ContentControlProperties? ContentControlProperties
    {
        get => _contentControlProperties;
        set => Set(ref _contentControlProperties, value);
    }

    /// <summary>Whether this run sits inside a content control.</summary>
    public bool IsContentControlRun => ContentControlProperties is not null;

    /// <summary>True when the caller assigned this run's text.</summary>
    public bool IsTextChanged => IsChanged(nameof(Text));

    /// <summary>True when this run or its formatting has pending changes.</summary>
    public bool HasRunChanges => HasChanges || _formatting.HasChanges;

    /// <summary>Clears the change record on this run and its formatting.</summary>
    public void AcceptAllChanges()
    {
        AcceptChanges();
        _formatting.AcceptChanges();
    }

    /// <summary>Creates an empty run.</summary>
    public FormattedRun() { }

    /// <summary>Creates a run with the given text.</summary>
    /// <param name="text">The run's text.</param>
    public FormattedRun(string text) => _text = text;

    /// <summary>Creates a run with the given text and formatting.</summary>
    /// <param name="text">The run's text.</param>
    /// <param name="formatting">The character formatting.</param>
    public FormattedRun(string text, RunFormatting formatting) => (_text, _formatting) = (text, formatting);

    /// <summary>
    /// Creates a copy holding part of this run's text, keeping everything else about it.
    /// </summary>
    /// <param name="text">The slice's text.</param>
    /// <returns>The slice.</returns>
    /// <remarks>
    /// Used when a run is split so a formatting change can apply to part of it. The slice keeps the
    /// content control properties and document property field of the run it came from — dropping
    /// them detached the run from its control, so a later edit through that control found nothing to
    /// change. The text is set without recording a change, because splitting a run does not alter
    /// the text the paragraph contains.
    /// </remarks>
    internal FormattedRun CloneWithText(string text)
    {
        var clone = Clone();
        clone._text = text;
        return clone;
    }

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public FormattedRun Clone()
    {
        var clone = new FormattedRun
        {
            _text = _text,
            _formatting = _formatting.Clone(),
            _isTab = _isTab,
            _isBreak = _isBreak,
            _breakType = _breakType,
            _documentPropertyField = _documentPropertyField?.Clone(),
            _contentControlProperties = _contentControlProperties?.Clone()
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
