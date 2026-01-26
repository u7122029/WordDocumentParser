using WordDocumentParser.Models.ContentControls;

namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Represents a run of text with its formatting.
/// </summary>
public class FormattedRun
{
    public string Text { get; set; } = string.Empty;
    public RunFormatting Formatting { get; set; } = new();
    public bool IsTab { get; set; }
    public bool IsBreak { get; set; }
    public string? BreakType { get; set; } // TextWrapping, Page, Column

    /// <summary>
    /// If this run represents a document property field, contains the field information
    /// </summary>
    public DocumentPropertyField? DocumentPropertyField { get; set; }

    /// <summary>
    /// Whether this run is a document property field
    /// </summary>
    public bool IsDocumentPropertyField => DocumentPropertyField is not null;

    /// <summary>
    /// If this run is part of an inline content control, contains the control properties
    /// </summary>
    public ContentControlProperties? ContentControlProperties { get; set; }

    /// <summary>
    /// Whether this run is part of a content control
    /// </summary>
    public bool IsContentControlRun => ContentControlProperties is not null;

    public FormattedRun() { }
    public FormattedRun(string text) => Text = text;
    public FormattedRun(string text, RunFormatting formatting) => (Text, Formatting) = (text, formatting);
}