using WordDocumentParser.Core;
using WordDocumentParser.Models.ContentControls;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser;

/// <summary>
/// Represents a node in the document tree structure.
/// Nodes form a hierarchy based on heading levels, with content nested under headings.
/// </summary>
public class DocumentNode(ContentType type)
{
    /// <summary>Unique identifier for this node</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>The content type of this node (Paragraph, Heading, Table, etc.)</summary>
    public ContentType Type { get; set; } = type;

    /// <summary>Heading level (1-9) or 0 for non-headings</summary>
    public int HeadingLevel { get; set; }

    /// <summary>Plain text content of this node</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Child nodes in document order</summary>
    public List<DocumentNode> Children { get; set; } = [];

    /// <summary>Parent node in the tree hierarchy</summary>
    public DocumentNode? Parent { get; set; }

    /// <summary>
    /// Additional metadata
    /// </summary>
    public Dictionary<string, object> Metadata { get; set; } = [];

    /// <summary>
    /// Formatted text runs that make up the text content with styling
    /// </summary>
    public List<FormattedRun> Runs { get; set; } = [];

    /// <summary>
    /// Paragraph-level formatting
    /// </summary>
    public ParagraphFormatting? ParagraphFormatting { get; set; }

    /// <summary>
    /// Original OpenXML content for exact round-trip (stores full paragraph/table XML)
    /// </summary>
    public string? OriginalXml { get; set; }

    /// <summary>
    /// Properties for content controls (SDT blocks). Only set for nodes that represent content controls.
    /// </summary>
    public ContentControlProperties? ContentControlProperties { get; set; }

    /// <summary>
    /// Returns true if this node is or contains a content control
    /// </summary>
    public bool IsContentControl => ContentControlProperties is not null ||
                                    (Metadata.TryGetValue("IsSdtContent", out var isSdt) && isSdt is true) ||
                                    (Metadata.TryGetValue("IsSdtBlock", out var isSdtBlock) && isSdtBlock is true);

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
    public DocumentNode(ContentType type, string text) : this(type) => Text = text;

    /// <summary>Creates a heading node with the specified level and text.</summary>
    public DocumentNode(ContentType type, int headingLevel, string text) : this(type, text) => HeadingLevel = headingLevel;

    /// <summary>
    /// Adds a child node and sets the parent reference
    /// </summary>
    public void AddChild(DocumentNode child)
    {
        child.Parent = this;
        Children.Add(child);
    }

    /// <summary>
    /// Gets the depth of this node in the tree
    /// </summary>
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
    /// Pretty prints the tree structure
    /// </summary>
    public string ToTreeString(int indent = 0)
    {
        var prefix = new string(' ', indent * 2);
        var typeLabel = Type == ContentType.Heading ? $"H{HeadingLevel}" : Type.ToString();
        var textPreview = Text.Length > 80 ? $"{Text[..77]}... ({GetTextWithMetadata()[..77]})" : $"{Text} ({GetTextWithMetadata()})";
        var result = $"{prefix}[{typeLabel}][{ParagraphFormatting?.StyleId}] {textPreview}\n";

        foreach (var child in Children)
        {
            result += child.ToTreeString(indent + 1);
        }

        return result;
    }

    /// <summary>Returns a short string representation of this node.</summary>
    public override string ToString()
    {
        var typeLabel = Type == ContentType.Heading ? $"Heading{HeadingLevel}" : Type.ToString();
        return $"{typeLabel}: {(Text.Length > 30 ? $"{Text[..77]}..." : Text)}";
    }
}
