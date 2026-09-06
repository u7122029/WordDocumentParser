using System.Text;
using WordDocumentParser.Core;
using WordDocumentParser.Models.ContentControls;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser;

/// <summary>
/// Represents a node in the document tree structure.
/// Nodes form a hierarchy based on heading levels, with content nested under headings.
/// </summary>
/// <remarks>
/// Assignments are tracked (see <see cref="TrackedModel"/>). A node parsed from a document reports
/// no changes, so the writer emits its <see cref="OriginalXml"/> untouched; once a caller assigns
/// <see cref="Text"/>, edits <see cref="Runs"/>, or changes formatting, the writer applies exactly
/// those edits to that original XML rather than regenerating the element, which would drop the
/// hyperlinks, fields, and bookmarks the model does not represent.
/// </remarks>
public class DocumentNode(ContentType type) : TrackedModel
{
    /// <summary>Metadata key marking a node whose original XML is a block-level SDT.</summary>
    public const string IsSdtBlockKey = "IsSdtBlock";

    /// <summary>Metadata key marking a node produced from the single paragraph inside an SDT.</summary>
    public const string IsSdtContentKey = "IsSdtContent";

    /// <summary>
    /// Metadata key marking a child node whose content is physically inside its parent's SDT XML,
    /// as opposed to a sibling the heading hierarchy merely nested beneath it.
    /// </summary>
    /// <remarks>
    /// The writer needs the distinction: content inside the SDT is already emitted with the parent's
    /// original XML and must not be written twice, while content merely nested under an SDT heading
    /// must still be written or the rest of the section disappears.
    /// </remarks>
    public const string IsInsideParentSdtKey = "IsInsideParentSdt";

    private string _text = string.Empty;
    private List<FormattedRun> _runs = [];
    private ParagraphFormatting? _paragraphFormatting;
    private string? _originalXml;
    private ContentControlProperties? _contentControlProperties;

    /// <summary>Unique identifier for this node</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>The content type of this node (Paragraph, Heading, Table, etc.)</summary>
    public ContentType Type { get; set; } = type;

    /// <summary>Heading level (1-9) or 0 for non-headings</summary>
    public int HeadingLevel { get; set; }

    /// <summary>
    /// Plain text content of this node. Assigning this marks the node as edited, so the writer
    /// updates the text in the node's original XML on the next save.
    /// </summary>
    public string Text { get => _text; set => Set(ref _text, value ?? string.Empty); }

    /// <summary>Child nodes in document order</summary>
    public List<DocumentNode> Children { get; set; } = [];

    /// <summary>Parent node in the tree hierarchy</summary>
    public DocumentNode? Parent { get; set; }

    /// <summary>
    /// Additional metadata
    /// </summary>
    public Dictionary<string, object> Metadata { get; set; } = [];

    /// <summary>
    /// Formatted text runs that make up the text content with styling.
    /// Mutating this list directly does not mark the node as edited; call
    /// <see cref="MarkRunsChanged"/> afterwards, or use the extension methods, which do so.
    /// </summary>
    public List<FormattedRun> Runs { get => _runs; set { Set(ref _runs, value ?? []); MarkRunsChanged(); } }

    /// <summary>
    /// Paragraph-level formatting
    /// </summary>
    public ParagraphFormatting? ParagraphFormatting
    {
        get => _paragraphFormatting;
        set => Set(ref _paragraphFormatting, value);
    }

    /// <summary>
    /// Original OpenXML content for exact round-trip (stores full paragraph/table XML)
    /// </summary>
    public string? OriginalXml { get => _originalXml; set => Set(ref _originalXml, value); }

    /// <summary>
    /// Properties for content controls (SDT blocks). Only set for nodes that represent content controls.
    /// </summary>
    public ContentControlProperties? ContentControlProperties
    {
        get => _contentControlProperties;
        set => Set(ref _contentControlProperties, value);
    }

    /// <summary>
    /// Returns true if this node is or contains a content control
    /// </summary>
    public bool IsContentControl => ContentControlProperties is not null ||
                                    (Metadata.TryGetValue(IsSdtContentKey, out var isSdt) && isSdt is true) ||
                                    (Metadata.TryGetValue(IsSdtBlockKey, out var isSdtBlock) && isSdtBlock is true);

    /// <summary>Returns true if this node's original XML is a block-level SDT.</summary>
    public bool IsSdtBlock => Metadata.TryGetValue(IsSdtBlockKey, out var value) && value is true;

    /// <summary>
    /// Returns true if this node's content is physically contained in its parent's SDT XML.
    /// </summary>
    public bool IsInsideParentSdt => Metadata.TryGetValue(IsInsideParentSdtKey, out var value) && value is true;

    #region Change tracking

    /// <summary>
    /// Records that this node's run list was replaced, reordered, added to, or removed from.
    /// </summary>
    public void MarkRunsChanged() => MarkChanged(nameof(Runs));

    /// <summary>True when a caller has assigned this node's text.</summary>
    public bool IsTextChanged => IsChanged(nameof(Text));

    /// <summary>True when this node's run list or any individual run has been edited.</summary>
    public bool IsRunsChanged => IsChanged(nameof(Runs)) || _runs.Exists(run => run.HasRunChanges);

    /// <summary>True when this node's paragraph formatting has been edited.</summary>
    public bool IsParagraphFormattingChanged =>
        IsChanged(nameof(ParagraphFormatting)) || _paragraphFormatting?.HasFormattingChanges is true;

    /// <summary>True when this node's content control properties have been edited.</summary>
    public bool IsContentControlChanged =>
        IsChanged(nameof(ContentControlProperties)) || _contentControlProperties?.HasChanges is true;

    /// <summary>
    /// True when this node or anything it owns — runs, formatting, content control properties, or
    /// the table data in its metadata — has been edited since it was parsed.
    /// </summary>
    public override bool HasChanges =>
        HasOwnChanges ||
        IsRunsChanged ||
        IsParagraphFormattingChanged ||
        IsContentControlChanged ||
        GetTableDataIfPresent()?.HasTableChanges is true;

    /// <summary>
    /// Clears the change record on this node and everything it owns, treating the current state as
    /// the unmodified baseline. Called by the parser once a node is fully populated.
    /// </summary>
    public void AcceptAllChanges()
    {
        AcceptChanges();

        foreach (var run in _runs)
        {
            run.AcceptAllChanges();
        }

        _paragraphFormatting?.AcceptAllChanges();
        _contentControlProperties?.AcceptChanges();
        GetTableDataIfPresent()?.AcceptAllChanges();

        foreach (var child in Children)
        {
            child.AcceptAllChanges();
        }
    }

    private Models.Tables.TableData? GetTableDataIfPresent() =>
        Metadata.TryGetValue("TableData", out var data) ? data as Models.Tables.TableData : null;

    #endregion

    /// <summary>
    /// Gets the plain text from formatted runs, or the Text property if no runs exist.
    /// Document property field values and content control values are included as their actual values.
    /// </summary>
    public string GetText() => Runs.Count > 0
        ? string.Concat(Runs.Select(r => r.IsTab ? "\t" : r.IsBreak ? " " : r.Text))
        : Text;

    /// <summary>
    /// Gets text with metadata annotations for document properties and content controls.
    /// Instead of showing just values, this shows metadata like property names, types, and current values.
    /// </summary>
    public string GetTextWithMetadata()
    {
        var textValue = GetText().Trim();

        // If this is a content control node with properties, always show metadata
        if (ContentControlProperties is not null)
        {
            var ccProps = ContentControlProperties;
            var identifier = !string.IsNullOrEmpty(ccProps.Alias) ? ccProps.Alias :
                            !string.IsNullOrEmpty(ccProps.Tag) ? ccProps.Tag :
                            ccProps.Id?.ToString() ?? "unnamed";

            // For document property content controls with data binding, show the property info
            if (ccProps.Type == ContentControlType.DocumentProperty && !string.IsNullOrEmpty(ccProps.DataBindingXPath))
            {
                var propName = DocumentPropertyHelpers.ExtractPropertyNameFromXPath(ccProps.DataBindingXPath);
                return $"[DocProperty:{propName}=\"{textValue}\"]";
            }

            return $"[ContentControl:{ccProps.Type} {identifier}=\"{textValue}\"]";
        }

        // Check for document property fields in runs
        if (Runs.Count == 0)
        {
            return Text;
        }

        var parts = new List<string>();

        // Group consecutive runs by their content control properties to avoid repeating metadata
        var i = 0;
        while (i < Runs.Count)
        {
            var run = Runs[i];

            if (run.IsTab)
            {
                parts.Add("\t");
                i++;
            }
            else if (run.IsBreak)
            {
                parts.Add(" ");
                i++;
            }
            else if (run.IsDocumentPropertyField && run.DocumentPropertyField is not null)
            {
                parts.Add(run.DocumentPropertyField.ToMetadataString());
                i++;
            }
            else if (run.IsContentControlRun && run.ContentControlProperties is not null)
            {
                // Collect all consecutive runs with the same content control
                var ccRuns = new List<FormattedRun> { run };
                var ccProps = run.ContentControlProperties;
                i++;
                while (i < Runs.Count &&
                       Runs[i].ContentControlProperties == ccProps &&
                       !Runs[i].IsTab && !Runs[i].IsBreak)
                {
                    ccRuns.Add(Runs[i]);
                    i++;
                }

                var ccText = string.Concat(ccRuns.Select(r => r.Text));
                var identifier = !string.IsNullOrEmpty(ccProps.Alias) ? ccProps.Alias :
                                !string.IsNullOrEmpty(ccProps.Tag) ? ccProps.Tag :
                                ccProps.Id?.ToString() ?? "unnamed";

                if (ccProps.Type == ContentControlType.DocumentProperty && !string.IsNullOrEmpty(ccProps.DataBindingXPath))
                {
                    var propName = DocumentPropertyHelpers.ExtractPropertyNameFromXPath(ccProps.DataBindingXPath);
                    parts.Add($"[DocProperty:{propName}=\"{ccText}\"]");
                }
                else
                {
                    parts.Add($"[ContentControl:{ccProps.Type} {identifier}=\"{ccText}\"]");
                }
            }
            else
            {
                parts.Add(run.Text);
                i++;
            }
        }

        return string.Concat(parts);
    }

    /// <summary>
    /// Returns true if this node has formatted runs
    /// </summary>
    public bool HasFormattedRuns => Runs.Count > 0;

    /// <summary>
    /// Returns true if this node contains any document property fields
    /// </summary>
    public bool HasDocumentPropertyFields => Runs.Any(r => r.IsDocumentPropertyField);

    /// <summary>Creates a node with the specified type and text content.</summary>
    /// <param name="type">The content type.</param>
    /// <param name="text">The node's text.</param>
    public DocumentNode(ContentType type, string text) : this(type) => _text = text;

    /// <summary>Creates a heading node with the specified level and text.</summary>
    /// <param name="type">The content type.</param>
    /// <param name="headingLevel">The heading level, 1 through 9.</param>
    /// <param name="text">The node's text.</param>
    public DocumentNode(ContentType type, int headingLevel, string text) : this(type, text) => HeadingLevel = headingLevel;

    /// <summary>
    /// Adds a child node and sets the parent reference
    /// </summary>
    /// <param name="child">The node to add.</param>
    public void AddChild(DocumentNode child)
    {
        child.Parent = this;
        Children.Add(child);
    }

    /// <summary>
    /// Gets the depth of this node in the tree
    /// </summary>
    /// <returns>The number of ancestors between this node and the root.</returns>
    public int GetDepth()
    {
        var depth = 0;
        var current = Parent;
        while (current is not null)
        {
            depth++;
            current = current.Parent;
        }
        return depth;
    }

    /// <summary>
    /// Pretty prints the tree structure.
    /// </summary>
    /// <param name="indent">Indentation level for this node.</param>
    /// <param name="previewLength">Maximum characters of text to show per node.</param>
    /// <returns>The rendered tree.</returns>
    /// <remarks>
    /// Rendered into a single buffer. Building the result by concatenating each subtree's string
    /// re-copied every ancestor's output once per descendant, which cost megabytes of allocation for
    /// a few hundred nodes.
    /// </remarks>
    public string ToTreeString(int indent = 0, int previewLength = 30)
    {
        var builder = new StringBuilder();
        AppendTreeString(builder, indent, previewLength);
        return builder.ToString();
    }

    private void AppendTreeString(StringBuilder builder, int indent, int previewLength)
    {
        var typeLabel = Type == ContentType.Heading ? $"H{HeadingLevel}" : Type.ToString();

        builder.Append(' ', indent * 2)
               .Append('[').Append(typeLabel).Append("][").Append(ParagraphFormatting?.StyleId).Append("] ");

        var metadataText = GetTextWithMetadata();
        if (Text.Length > previewLength)
        {
            builder.Append(Truncate(Text, previewLength)).Append("... (")
                   .Append(Truncate(metadataText, previewLength)).Append(')');
        }
        else
        {
            builder.Append(Text).Append(" (").Append(metadataText).Append(')');
        }

        builder.Append('\n');

        foreach (var child in Children)
        {
            child.AppendTreeString(builder, indent + 1, previewLength);
        }
    }

    /// <summary>
    /// Truncates to at most <paramref name="length"/> characters without slicing past the end.
    /// </summary>
    private static string Truncate(string value, int length) =>
        length <= 0 ? string.Empty :
        value.Length <= length ? value : value[..length];

    /// <summary>Returns a short string representation of this node.</summary>
    /// <returns>The node's type and a 30-character text preview.</returns>
    public override string ToString() => ToString(30);

    /// <summary>
    /// Returns a short string representation of this node, with adjustable preview length.
    /// </summary>
    /// <param name="previewLength">Maximum characters of text to show.</param>
    /// <returns>The node's type and a text preview.</returns>
    public string ToString(int previewLength)
    {
        var typeLabel = Type == ContentType.Heading ? $"Heading{HeadingLevel}" : Type.ToString();
        return $"{typeLabel}: {(Text.Length > previewLength ? $"{Truncate(Text, previewLength)}..." : Text)}";
    }
}
