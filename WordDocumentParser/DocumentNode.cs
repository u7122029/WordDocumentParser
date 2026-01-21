using System;
using System.Collections.Generic;
using System.Linq;

namespace WordDocumentParser
{
    /// <summary>
    /// Represents the type of content in a document node
    /// </summary>
    public enum ContentType
    {
        Document,
        Heading,
        Paragraph,
        Table,
        Image,
        List,
        ListItem,
        HyperlinkText,
        TextRun
    }

    /// <summary>
    /// Represents a node in the document tree structure
    /// </summary>
    public class DocumentNode
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public ContentType Type { get; set; }
        public int HeadingLevel { get; set; } // 0 for non-headings, 1-9 for H1-H9
        public string Text { get; set; } = string.Empty;
        public List<DocumentNode> Children { get; set; } = new List<DocumentNode>();
        public DocumentNode? Parent { get; set; }

        // Additional metadata
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();

        // Formatting properties for round-trip fidelity
        /// <summary>
        /// Formatted text runs that make up the text content with styling
        /// </summary>
        public List<FormattedRun> Runs { get; set; } = new List<FormattedRun>();

        /// <summary>
        /// Paragraph-level formatting
        /// </summary>
        public ParagraphFormatting? ParagraphFormatting { get; set; }

        /// <summary>
        /// Original document package data for round-trip fidelity (only set on root Document node)
        /// </summary>
        public DocumentPackageData? PackageData { get; set; }

        /// <summary>
        /// Original OpenXML content for exact round-trip (stores full paragraph/table XML)
        /// </summary>
        public string? OriginalXml { get; set; }

        /// <summary>
        /// Gets the plain text from formatted runs, or the Text property if no runs exist
        /// </summary>
        public string GetText()
        {
            if (Runs.Count > 0)
            {
                return string.Join("", Runs.Select(r => r.IsTab ? "\t" : (r.IsBreak ? " " : r.Text)));
            }
            return Text;
        }

        /// <summary>
        /// Returns true if this node has formatted runs
        /// </summary>
        public bool HasFormattedRuns => Runs.Count > 0;

        public DocumentNode(ContentType type)
        {
            Type = type;
        }

        public DocumentNode(ContentType type, string text) : this(type)
        {
            Text = text;
        }

        public DocumentNode(ContentType type, int headingLevel, string text) : this(type, text)
        {
            HeadingLevel = headingLevel;
        }

        /// <summary>
        /// Adds a child node and sets the parent reference
        /// </summary>
        public void AddChild(DocumentNode child)
        {
            child.Parent = this;
            Children.Add(child);
        }

        /// <summary>
        /// Gets all descendant nodes (recursive)
        /// </summary>
        public IEnumerable<DocumentNode> GetDescendants()
        {
            foreach (var child in Children)
            {
                yield return child;
                foreach (var descendant in child.GetDescendants())
                {
                    yield return descendant;
                }
            }
        }

        /// <summary>
        /// Gets the depth of this node in the tree
        /// </summary>
        public int GetDepth()
        {
            int depth = 0;
            var current = Parent;
            while (current != null)
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
            var textPreview = Text.Length > 50 ? Text.Substring(0, 47) + "..." : Text;
            var result = $"{prefix}[{typeLabel}][{ParagraphFormatting?.StyleId}] {textPreview}\n";

            foreach (var child in Children)
            {
                result += child.ToTreeString(indent + 1);
            }

            return result;
        }

        public override string ToString()
        {
            var typeLabel = Type == ContentType.Heading ? $"Heading{HeadingLevel}" : Type.ToString();
            return $"{typeLabel}: {(Text.Length > 30 ? Text.Substring(0, 27) + "..." : Text)}";
        }
    }
}
