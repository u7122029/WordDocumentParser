using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Images;
using WordDocumentParser.Models.Package;
using WordDocumentParser.Models.Tables;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for merging and concatenating Word documents.
/// </summary>
public static class DocumentMergeExtensions
{
    #region Public API

    /// <summary>
    /// Appends the content of another document to this document.
    /// Images, hyperlinks, and other resources are merged automatically.
    /// </summary>
    /// <param name="target">The document to append content to</param>
    /// <param name="source">The document whose content will be appended</param>
    /// <param name="addPageBreak">If true, adds a page break before the appended content</param>
    /// <returns>The target document (for method chaining)</returns>
    public static WordDocument AppendDocument(this WordDocument target, WordDocument source, bool addPageBreak = true)
    {
        var mapping = MergeResources(target, source);

        // Optionally add a page break separator
        if (addPageBreak)
        {
            var pageBreakNode = CreatePageBreakNode();
            target.Root.AddChild(pageBreakNode);
        }

        // Clone and append all top-level nodes from source
        foreach (var child in source.Root.Children)
        {
            var clonedNode = CloneNode(child, mapping);
            target.Root.AddChild(clonedNode);
        }

        return target;
    }

    /// <summary>
    /// Appends the content of multiple documents to this document.
    /// </summary>
    /// <param name="target">The document to append content to</param>
    /// <param name="sources">The documents whose content will be appended (in order)</param>
    /// <param name="addPageBreaks">If true, adds page breaks between documents</param>
    /// <returns>The target document (for method chaining)</returns>
    public static WordDocument AppendDocuments(this WordDocument target, IEnumerable<WordDocument> sources, bool addPageBreaks = true)
    {
        foreach (var source in sources)
        {
            target.AppendDocument(source, addPageBreaks);
        }
        return target;
    }

    /// <summary>
    /// Creates a new document by concatenating multiple documents together.
    /// The first document's styles, theme, and settings are preserved.
    /// </summary>
    /// <param name="documents">The documents to concatenate (in order)</param>
    /// <param name="addPageBreaks">If true, adds page breaks between documents</param>
    /// <returns>A new document containing all content from the input documents</returns>
    public static WordDocument ConcatenateDocuments(IEnumerable<WordDocument> documents, bool addPageBreaks = true)
    {
        var docList = documents.ToList();
        if (docList.Count == 0)
        {
            return new WordDocument();
        }

        // Use the first document as the base (clone it to avoid modifying the original)
        var result = CloneDocument(docList[0]);

        // Append remaining documents
        for (int i = 1; i < docList.Count; i++)
        {
            result.AppendDocument(docList[i], addPageBreaks);
        }

        return result;
    }

    /// <summary>
    /// Creates an independent copy of a document: the content tree, the package data, and the
    /// mutable models hanging off both.
    /// </summary>
    /// <param name="source">The document to clone</param>
    /// <returns>A new document that shares no mutable state with the source</returns>
    /// <remarks>
    /// Editing the copy leaves the source untouched: the mutable models behind the tree, including
    /// the table data held in node metadata, are copied rather than shared.
    /// </remarks>
    public static WordDocument CloneDocument(WordDocument source)
    {
        // Property edits live in a dictionary that is only serialised on save, so flush them first
        // or the copy inherits the package's stale XML and loses everything set since parsing.
        source.SyncCustomPropertiesToXml();

        var clonedRoot = CloneNode(source.Root, ResourceMapping.Empty);

        return new WordDocument(clonedRoot)
        {
            FileName = source.FileName,
            PackageData = ClonePackageData(source.PackageData)
        };
    }

    /// <summary>
    /// Gets statistics about a document merge operation.
    /// </summary>
    public static MergeStatistics GetMergeStatistics(this WordDocument target, WordDocument source)
    {
        return new MergeStatistics
        {
            TargetNodeCount = CountAllNodes(target.Root),
            SourceNodeCount = CountAllNodes(source.Root),
            TargetImageCount = target.PackageData.Images.Count,
            SourceImageCount = source.PackageData.Images.Count,
            TargetHyperlinkCount = target.PackageData.HyperlinkRelationships.Count,
            SourceHyperlinkCount = source.PackageData.HyperlinkRelationships.Count
        };
    }

    #endregion

    #region Section Extraction

    /// <summary>
    /// Extracts a section from a document by heading text.
    /// Returns the heading node and all its nested content (up to the next heading of same or higher level).
    /// </summary>
    /// <param name="document">The source document</param>
    /// <param name="headingText">The heading text to find (case-insensitive partial match)</param>
    /// <param name="includeNestedHeadings">If true, includes sub-headings; if false, stops at next heading</param>
    /// <returns>List of nodes representing the section, or empty if heading not found</returns>
    public static List<DocumentNode> ExtractSection(this WordDocument document, string headingText, bool includeNestedHeadings = true)
    {
        return document.Root.ExtractSection(headingText, includeNestedHeadings);
    }

    /// <summary>
    /// Extracts a section from a node tree by heading text.
    /// </summary>
    /// <param name="root">The node to search under</param>
    /// <param name="headingText">The heading text to find (case-insensitive partial match)</param>
    /// <param name="includeNestedHeadings">If true, includes sub-headings; if false, excludes them</param>
    /// <returns>The heading and its content, or an empty list if the heading was not found</returns>
    /// <remarks>
    /// <para>
    /// The heading is found anywhere in the tree, and its content is whatever the tree nests beneath
    /// it. The tree is built by heading hierarchy, so an H2 is a child of its H1 rather than a
    /// sibling, and a flat scan of the root's children would never reach it.
    /// </para>
    /// <para>
    /// With <paramref name="includeNestedHeadings"/> false, sub-headings and their content are left
    /// out; only the heading's own direct content comes back, as a detached copy.
    /// </para>
    /// </remarks>
    public static List<DocumentNode> ExtractSection(this DocumentNode root, string headingText, bool includeNestedHeadings = true)
    {
        var heading = root.FindSectionHeading(headingText);
        if (heading is null)
            return [];

        // The tree nests a section's content under its heading, so the heading node is the section.
        if (includeNestedHeadings)
            return [heading];

        // Excluding sub-headings means returning a copy: the sub-headings are children of the live
        // heading, so leaving them out of the list would not actually leave them out of the section.
        var trimmed = CloneNode(heading, ResourceMapping.Empty);
        trimmed.Children.RemoveAll(child => child.Type == ContentType.Heading);
        return [trimmed];
    }

    /// <summary>
    /// Finds the heading that starts a section, anywhere in the tree.
    /// </summary>
    /// <param name="root">The node to search under.</param>
    /// <param name="headingText">The heading text to find (case-insensitive partial match).</param>
    /// <returns>The heading node, or null when no heading matches.</returns>
    internal static DocumentNode? FindSectionHeading(this DocumentNode root, string headingText) =>
        root.FindAll(n =>
            n.Type == ContentType.Heading &&
            n.GetText().Contains(headingText, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

    /// <summary>
    /// Extracts nodes from a document that match a predicate.
    /// </summary>
    /// <param name="document">The source document</param>
    /// <param name="predicate">Function to determine which nodes to include</param>
    /// <param name="includeChildren">If true, includes children of matching nodes</param>
    /// <returns>List of matching nodes</returns>
    public static List<DocumentNode> ExtractNodes(this WordDocument document, Func<DocumentNode, bool> predicate, bool includeChildren = true)
    {
        var result = new List<DocumentNode>();

        foreach (var node in document.Root.Children)
        {
            if (predicate(node))
            {
                result.Add(node);
            }
            else if (includeChildren)
            {
                // Check children recursively
                result.AddRange(ExtractNodesRecursive(node, predicate));
            }
        }

        return result;
    }

    private static List<DocumentNode> ExtractNodesRecursive(DocumentNode node, Func<DocumentNode, bool> predicate)
    {
        var result = new List<DocumentNode>();

        foreach (var child in node.Children)
        {
            if (predicate(child))
            {
                result.Add(child);
            }
            result.AddRange(ExtractNodesRecursive(child, predicate));
        }

        return result;
    }

    /// <summary>
    /// Extracts a range of top-level nodes by index.
    /// </summary>
    /// <param name="document">The source document</param>
    /// <param name="startIndex">Starting index (0-based)</param>
    /// <param name="count">Number of nodes to extract</param>
    /// <returns>List of nodes in the range</returns>
    public static List<DocumentNode> ExtractNodeRange(this WordDocument document, int startIndex, int count)
    {
        return document.Root.Children
            .Skip(startIndex)
            .Take(count)
            .ToList();
    }

    /// <summary>
    /// Extracts all tables from a document.
    /// </summary>
    public static List<DocumentNode> ExtractTables(this WordDocument document)
    {
        return document.Root.FindAll(n => n.Type == ContentType.Table).ToList();
    }

    /// <summary>
    /// Extracts all headings at a specific level from a document.
    /// </summary>
    public static List<DocumentNode> ExtractHeadingsAtLevel(this WordDocument document, int level)
    {
        return document.Root.FindAll(n => n.Type == ContentType.Heading && n.HeadingLevel == level).ToList();
    }

    #endregion

    #region Node Insertion

    /// <summary>
    /// Inserts nodes from a source document after a specific node in the target document.
    /// Resources (images, hyperlinks) are automatically merged.
    /// </summary>
    /// <param name="target">The target document</param>
    /// <param name="afterNode">The node after which to insert</param>
    /// <param name="sourceNodes">The nodes to insert</param>
    /// <param name="source">The source document (for resource merging)</param>
    /// <returns>The target document (for method chaining)</returns>
    public static WordDocument InsertNodesAfter(
        this WordDocument target,
        DocumentNode afterNode,
        IEnumerable<DocumentNode> sourceNodes,
        WordDocument source)
    {
        var parent = afterNode.Parent;
        if (parent == null)
            throw new ArgumentException("The target node must have a parent", nameof(afterNode));

        var index = parent.Children.IndexOf(afterNode);
        if (index < 0)
            throw new ArgumentException("The target node was not found in its parent's children", nameof(afterNode));

        return target.InsertNodesAtIndex(parent, index + 1, sourceNodes, source);
    }

    /// <summary>
    /// Inserts nodes from a source document before a specific node in the target document.
    /// Resources (images, hyperlinks) are automatically merged.
    /// </summary>
    /// <param name="target">The target document</param>
    /// <param name="beforeNode">The node before which to insert</param>
    /// <param name="sourceNodes">The nodes to insert</param>
    /// <param name="source">The source document (for resource merging)</param>
    /// <returns>The target document (for method chaining)</returns>
    public static WordDocument InsertNodesBefore(
        this WordDocument target,
        DocumentNode beforeNode,
        IEnumerable<DocumentNode> sourceNodes,
        WordDocument source)
    {
        var parent = beforeNode.Parent;
        if (parent == null)
            throw new ArgumentException("The target node must have a parent", nameof(beforeNode));

        var index = parent.Children.IndexOf(beforeNode);
        if (index < 0)
            throw new ArgumentException("The target node was not found in its parent's children", nameof(beforeNode));

        return target.InsertNodesAtIndex(parent, index, sourceNodes, source);
    }

    /// <summary>
    /// Inserts nodes from a source document at a specific index in the target document's root.
    /// Resources (images, hyperlinks) are automatically merged.
    /// </summary>
    /// <param name="target">The target document</param>
    /// <param name="index">The index at which to insert (0-based)</param>
    /// <param name="sourceNodes">The nodes to insert</param>
    /// <param name="source">The source document (for resource merging)</param>
    /// <returns>The target document (for method chaining)</returns>
    public static WordDocument InsertNodesAt(
        this WordDocument target,
        int index,
        IEnumerable<DocumentNode> sourceNodes,
        WordDocument source)
    {
        return target.InsertNodesAtIndex(target.Root, index, sourceNodes, source);
    }

    /// <summary>
    /// Inserts nodes at a specific index within a parent node.
    /// </summary>
    private static WordDocument InsertNodesAtIndex(
        this WordDocument target,
        DocumentNode parent,
        int index,
        IEnumerable<DocumentNode> sourceNodes,
        WordDocument source)
    {
        var mapping = MergeResources(target, source);

        var nodesToInsert = sourceNodes.ToList();
        var insertIndex = Math.Clamp(index, 0, parent.Children.Count);

        for (var i = 0; i < nodesToInsert.Count; i++)
        {
            var clonedNode = CloneNode(nodesToInsert[i], mapping);
            clonedNode.Parent = parent;
            parent.Children.Insert(insertIndex + i, clonedNode);
        }

        return target;
    }

    /// <summary>
    /// Inserts a section from a source document after a heading in the target document.
    /// </summary>
    /// <param name="target">The target document</param>
    /// <param name="afterHeadingText">Text of the heading after which to insert</param>
    /// <param name="source">The source document</param>
    /// <param name="sectionHeadingText">Text of the section heading to copy from source</param>
    /// <param name="includeNestedHeadings">Whether to include sub-headings from the source section</param>
    /// <returns>The target document (for method chaining)</returns>
    public static WordDocument InsertSectionAfterHeading(
        this WordDocument target,
        string afterHeadingText,
        WordDocument source,
        string sectionHeadingText,
        bool includeNestedHeadings = true)
    {
        // Find the target heading
        var targetHeading = target.Root.FindAll(n =>
            n.Type == ContentType.Heading &&
            n.GetText().Contains(afterHeadingText, StringComparison.OrdinalIgnoreCase))
            .FirstOrDefault();

        if (targetHeading == null)
            throw new ArgumentException($"Heading '{afterHeadingText}' not found in target document", nameof(afterHeadingText));

        // Extract the section from source
        var sectionNodes = source.ExtractSection(sectionHeadingText, includeNestedHeadings);
        if (sectionNodes.Count == 0)
            throw new ArgumentException($"Section '{sectionHeadingText}' not found in source document", nameof(sectionHeadingText));

        // Find the last node under the target heading (to insert after it)
        var targetHeadingLevel = targetHeading.HeadingLevel;
        var parent = targetHeading.Parent;
        if (parent == null)
            throw new InvalidOperationException("Target heading has no parent");

        var headingIndex = parent.Children.IndexOf(targetHeading);

        // Find where the target heading's content ends
        int insertIndex = headingIndex + 1;
        for (int i = headingIndex + 1; i < parent.Children.Count; i++)
        {
            var node = parent.Children[i];
            if (node.Type == ContentType.Heading && node.HeadingLevel <= targetHeadingLevel)
            {
                break;
            }
            insertIndex = i + 1;
        }

        return target.InsertNodesAtIndex(parent, insertIndex, sectionNodes, source);
    }

    /// <summary>
    /// Replaces a section in the target document with a section from the source document.
    /// </summary>
    /// <param name="target">The target document</param>
    /// <param name="targetHeadingText">Text of the section heading to replace</param>
    /// <param name="source">The source document</param>
    /// <param name="sourceHeadingText">Text of the section heading to copy from source</param>
    /// <param name="includeNestedHeadings">Whether to include sub-headings</param>
    /// <returns>The target document (for method chaining)</returns>
    /// <remarks>
    /// Both sections are located before anything is removed, so a missing source section leaves the
    /// target exactly as it was rather than destroying the section it was asked to replace.
    /// </remarks>
    public static WordDocument ReplaceSection(
        this WordDocument target,
        string targetHeadingText,
        WordDocument source,
        string sourceHeadingText,
        bool includeNestedHeadings = true)
    {
        // The live heading, so it can be detached from its parent below.
        var targetHeading = target.Root.FindSectionHeading(targetHeadingText)
                            ?? throw new ArgumentException(
                                $"Section '{targetHeadingText}' not found in target document", nameof(targetHeadingText));

        // Validate the replacement before touching the target.
        var sourceSection = source.ExtractSection(sourceHeadingText, includeNestedHeadings);
        if (sourceSection.Count == 0)
            throw new ArgumentException($"Section '{sourceHeadingText}' not found in source document", nameof(sourceHeadingText));

        var parent = targetHeading.Parent
                     ?? throw new InvalidOperationException("Section has no parent");

        var insertIndex = parent.Children.IndexOf(targetHeading);
        parent.Children.Remove(targetHeading);

        return target.InsertNodesAtIndex(parent, insertIndex, sourceSection, source);
    }

    #endregion

    #region Resource Merging

    /// <summary>
    /// How a source document's resources were renamed to fit into a target document.
    /// </summary>
    internal sealed class ResourceMapping
    {
        /// <summary>A mapping that renames nothing, for cloning within one document.</summary>
        public static ResourceMapping Empty { get; } = new();

        /// <summary>Source relationship ID to target relationship ID.</summary>
        public Dictionary<string, string> Relationships { get; } = new(StringComparer.Ordinal);

        /// <summary>Source numbering ID to target numbering ID.</summary>
        public Dictionary<int, int> Numbering { get; } = [];

        /// <summary>True when nothing needs rewriting.</summary>
        public bool IsEmpty => Relationships.Count == 0 && Numbering.Count == 0;
    }

    /// <summary>
    /// Copies the source document's images, hyperlinks, and numbering definitions into the target.
    /// </summary>
    /// <remarks>
    /// All relationship IDs come from one allocator over the target's whole relationship namespace.
    /// Allocating per resource kind handed out IDs another kind already held: merging images into a
    /// document whose hyperlink was <c>rId1000</c> reassigned that ID and broke the link.
    /// </remarks>
    private static ResourceMapping MergeResources(WordDocument target, WordDocument source)
    {
        var mapping = new ResourceMapping();
        var packageData = target.PackageData;

        foreach (var (oldId, imageData) in source.PackageData.Images)
        {
            var newId = packageData.AllocateRelationshipId();

            packageData.Images[newId] = new ImagePartData
            {
                ContentType = imageData.ContentType,
                // Media bytes are shared and treated as immutable, so the copy references them.
                Data = imageData.Data,
                OriginalRelationshipId = newId,
                OriginalUri = imageData.OriginalUri
            };

            mapping.Relationships[oldId] = newId;
        }

        foreach (var (oldId, hyperlinkData) in source.PackageData.HyperlinkRelationships)
        {
            var newId = packageData.AllocateRelationshipId();

            packageData.HyperlinkRelationships[newId] = new HyperlinkRelationshipData
            {
                Url = hyperlinkData.Url,
                IsExternal = hyperlinkData.IsExternal
            };

            mapping.Relationships[oldId] = newId;
        }

        foreach (var (oldNumId, newNumId) in NumberingMerge.Merge(target, source))
        {
            mapping.Numbering[oldNumId] = newNumId;
        }

        return mapping;
    }

    #endregion

    #region Node Cloning

    /// <summary>
    /// Deep clones nodes from a source document, updating resource references for use in the target document.
    /// Call this when you need to manually clone nodes with proper resource handling.
    /// </summary>
    /// <param name="target">The target document (resources will be merged into this)</param>
    /// <param name="source">The source document</param>
    /// <param name="nodes">The nodes to clone</param>
    /// <returns>List of cloned nodes ready to be added to the target document</returns>
    public static List<DocumentNode> CloneNodesForDocument(
        this WordDocument target,
        WordDocument source,
        IEnumerable<DocumentNode> nodes)
    {
        var mapping = MergeResources(target, source);
        return nodes.Select(node => CloneNode(node, mapping)).ToList();
    }

    /// <summary>
    /// Deep clones a single node from a source document, updating resource references for use in the target document.
    /// </summary>
    /// <param name="target">The target document (resources will be merged into this)</param>
    /// <param name="source">The source document</param>
    /// <param name="node">The node to clone</param>
    /// <returns>The cloned node, ready to be added to the target document</returns>
    public static DocumentNode CloneNodeForDocument(
        this WordDocument target,
        WordDocument source,
        DocumentNode node)
        => CloneNode(node, MergeResources(target, source));

    /// <summary>
    /// Deep clones a document node and everything hanging off it, rewriting resource references.
    /// </summary>
    /// <remarks>
    /// Every mutable model is copied, including the table data held in metadata. Copying metadata
    /// values by reference meant a concatenated document shared its tables with the document it was
    /// built from, so editing one edited both.
    /// </remarks>
    internal static DocumentNode CloneNode(DocumentNode source, ResourceMapping mapping)
    {
        var clone = new DocumentNode(source.Type)
        {
            Id = Guid.NewGuid().ToString(),
            HeadingLevel = source.HeadingLevel,
            Text = source.Text,
            ParagraphFormatting = source.ParagraphFormatting?.Clone(),
            OriginalXml = RewriteReferences(source.OriginalXml, mapping),
            ContentControlProperties = source.ContentControlProperties?.Clone()
        };

        foreach (var (key, value) in source.Metadata)
        {
            clone.Metadata[key] = CloneMetadataValue(value, mapping);
        }

        foreach (var run in source.Runs)
        {
            clone.Runs.Add(run.Clone());
        }

        if (clone.ParagraphFormatting is { NumberingId: { } numberingId } formatting &&
            mapping.Numbering.TryGetValue(numberingId, out var newNumberingId))
        {
            formatting.NumberingId = newNumberingId;
        }

        // Skip Image children when the parent has OriginalXml: the image is already in that XML with
        // its rewritten relationship ID, and writing the child too would stack a duplicate on top.
        var hasOriginalXml = !string.IsNullOrEmpty(source.OriginalXml);

        foreach (var child in source.Children)
        {
            if (hasOriginalXml && child.Type == ContentType.Image)
            {
                continue;
            }

            clone.AddChild(CloneNode(child, mapping));
        }

        // The clone reproduces the source exactly, including whatever edits were pending on it, so
        // it inherits the source's change state rather than presenting itself as freshly parsed.
        if (!source.HasChanges)
        {
            clone.AcceptAllChanges();
        }

        return clone;
    }

    /// <summary>
    /// Copies a metadata value, deep-copying the mutable models the library stores there.
    /// </summary>
    private static object CloneMetadataValue(object value, ResourceMapping mapping) => value switch
    {
        TableData tableData => CloneTableData(tableData, mapping),
        List<HyperlinkData> hyperlinks => hyperlinks.Select(link => CloneHyperlink(link, mapping)).ToList(),
        ImageData imageData => CloneImageData(imageData, mapping),
        _ => value
    };

    private static TableData CloneTableData(TableData source, ResourceMapping mapping)
    {
        var clone = new TableData
        {
            ColumnCount = source.ColumnCount,
            Formatting = source.Formatting?.Clone()
        };

        foreach (var row in source.Rows)
        {
            var clonedRow = new TableRow
            {
                RowIndex = row.RowIndex,
                IsHeader = row.IsHeader,
                Formatting = row.Formatting?.Clone()
            };

            foreach (var cell in row.Cells)
            {
                var clonedCell = new TableCell
                {
                    RowIndex = cell.RowIndex,
                    ColumnIndex = cell.ColumnIndex,
                    RowSpan = cell.RowSpan,
                    ColSpan = cell.ColSpan,
                    Formatting = cell.Formatting?.Clone()
                };

                foreach (var content in cell.Content)
                {
                    clonedCell.Content.Add(CloneNode(content, mapping));
                }

                clonedRow.Cells.Add(clonedCell);
            }

            clone.Rows.Add(clonedRow);
        }

        if (!source.HasTableChanges)
        {
            clone.AcceptAllChanges();
        }

        return clone;
    }

    private static HyperlinkData CloneHyperlink(HyperlinkData source, ResourceMapping mapping) => new()
    {
        Text = source.Text,
        RelationshipId = source.RelationshipId is { } id && mapping.Relationships.TryGetValue(id, out var newId)
            ? newId
            : source.RelationshipId,
        Url = source.Url,
        Anchor = source.Anchor,
        Tooltip = source.Tooltip,
        Runs = source.Runs.Select(run => run.Clone()).ToList()
    };

    private static ImageData CloneImageData(ImageData source, ResourceMapping mapping) => new()
    {
        Id = mapping.Relationships.TryGetValue(source.Id, out var newId) ? newId : source.Id,
        Name = source.Name,
        ContentType = source.ContentType,
        // Media bytes are shared and treated as immutable.
        Data = source.Data,
        WidthInches = source.WidthInches,
        HeightInches = source.HeightInches,
        AltText = source.AltText,
        Description = source.Description,
        WidthEmu = source.WidthEmu,
        HeightEmu = source.HeightEmu,
        Formatting = source.Formatting
    };

    /// <summary>
    /// Rewrites the relationship and numbering references in a node's XML.
    /// </summary>
    private static string? RewriteReferences(string? originalXml, ResourceMapping mapping)
    {
        if (string.IsNullOrEmpty(originalXml) || mapping.IsEmpty)
            return originalXml;

        var result = originalXml;

        foreach (var (oldId, newId) in mapping.Relationships)
        {
            result = result
                .Replace($"r:embed=\"{oldId}\"", $"r:embed=\"{newId}\"")
                .Replace($"r:link=\"{oldId}\"", $"r:link=\"{newId}\"")
                .Replace($"r:id=\"{oldId}\"", $"r:id=\"{newId}\"");
        }

        return NumberingMerge.ApplyMapping(result, mapping.Numbering);
    }

    /// <summary>
    /// Creates a deep copy of DocumentPackageData.
    /// </summary>
    private static DocumentPackageData ClonePackageData(DocumentPackageData source)
    {
        var clone = new DocumentPackageData
        {
            StylesXml = source.StylesXml,
            ThemeXml = source.ThemeXml,
            FontTableXml = source.FontTableXml,
            NumberingXml = source.NumberingXml,
            SettingsXml = source.SettingsXml,
            WebSettingsXml = source.WebSettingsXml,
            FootnotesXml = source.FootnotesXml,
            EndnotesXml = source.EndnotesXml,
            CustomPropertiesXml = source.CustomPropertiesXml,
            CorePropertiesXml = source.CorePropertiesXml,
            AppPropertiesXml = source.AppPropertiesXml,
            GlossaryDocumentXml = source.GlossaryDocumentXml,
            GlossaryStylesXml = source.GlossaryStylesXml,
            GlossaryFontTableXml = source.GlossaryFontTableXml,
            OriginalDocumentXml = source.OriginalDocumentXml,

            // The source package is immutable once captured, so the copy shares the same bytes and
            // keeps the ability to save with everything the model does not represent intact.
            OriginalPackageBytes = source.OriginalPackageBytes,
            Baseline = source.Baseline,
            KnownRelationshipIds = new HashSet<string>(source.KnownRelationshipIds, StringComparer.Ordinal)
        };

        // Clone core and extended properties, carrying their pending changes so a removal made
        // before cloning is not lost in the copy.
        clone.CoreProperties = source.CoreProperties?.Clone();
        clone.ExtendedProperties = source.ExtendedProperties?.Clone();

        // Clone dictionaries
        foreach (var kvp in source.Headers)
            clone.Headers[kvp.Key] = kvp.Value;

        foreach (var kvp in source.Footers)
            clone.Footers[kvp.Key] = kvp.Value;

        foreach (var kvp in source.Images)
        {
            clone.Images[kvp.Key] = new ImagePartData
            {
                ContentType = kvp.Value.ContentType,
                Data = kvp.Value.Data,
                OriginalRelationshipId = kvp.Value.OriginalRelationshipId,
                OriginalUri = kvp.Value.OriginalUri
            };
        }

        foreach (var kvp in source.HyperlinkRelationships)
        {
            clone.HyperlinkRelationships[kvp.Key] = new HyperlinkRelationshipData
            {
                Url = kvp.Value.Url,
                IsExternal = kvp.Value.IsExternal
            };
        }

        foreach (var kvp in source.CustomXmlParts)
        {
            clone.CustomXmlParts[kvp.Key] = new CustomXmlPartData
            {
                XmlContent = kvp.Value.XmlContent,
                PropertiesXml = kvp.Value.PropertiesXml
            };
        }

        foreach (var kvp in source.HeaderImages)
        {
            clone.HeaderImages[kvp.Key] = new Dictionary<string, ImagePartData>();
            foreach (var imgKvp in kvp.Value)
            {
                clone.HeaderImages[kvp.Key][imgKvp.Key] = new ImagePartData
                {
                    ContentType = imgKvp.Value.ContentType,
                    Data = imgKvp.Value.Data,
                    OriginalRelationshipId = imgKvp.Value.OriginalRelationshipId,
                    OriginalUri = imgKvp.Value.OriginalUri
                };
            }
        }

        foreach (var kvp in source.FooterImages)
        {
            clone.FooterImages[kvp.Key] = new Dictionary<string, ImagePartData>();
            foreach (var imgKvp in kvp.Value)
            {
                clone.FooterImages[kvp.Key][imgKvp.Key] = new ImagePartData
                {
                    ContentType = imgKvp.Value.ContentType,
                    Data = imgKvp.Value.Data,
                    OriginalRelationshipId = imgKvp.Value.OriginalRelationshipId,
                    OriginalUri = imgKvp.Value.OriginalUri
                };
            }
        }

        foreach (var kvp in source.GlossaryImages)
        {
            clone.GlossaryImages[kvp.Key] = new ImagePartData
            {
                ContentType = kvp.Value.ContentType,
                Data = kvp.Value.Data,
                OriginalRelationshipId = kvp.Value.OriginalRelationshipId,
                OriginalUri = kvp.Value.OriginalUri
            };
        }

        clone.SectionPropertiesXml.AddRange(source.SectionPropertiesXml);

        return clone;
    }

    #endregion

    #region Helpers

    /// <summary>
    /// Creates a page break node to separate documents.
    /// </summary>
    private static DocumentNode CreatePageBreakNode()
    {
        return new DocumentNode(ContentType.Paragraph)
        {
            Text = "",
            OriginalXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:r><w:br w:type=""page""/></w:r></w:p>"
        };
    }

    /// <summary>
    /// Counts all nodes in a document tree.
    /// </summary>
    private static int CountAllNodes(DocumentNode root)
    {
        var count = 1;
        foreach (var child in root.Children)
        {
            count += CountAllNodes(child);
        }
        return count;
    }

    #endregion
}

/// <summary>
/// Statistics about a document merge operation.
/// </summary>
public class MergeStatistics
{
    /// <summary>Total nodes in the target document before merge</summary>
    public int TargetNodeCount { get; init; }

    /// <summary>Total nodes in the source document</summary>
    public int SourceNodeCount { get; init; }

    /// <summary>Images in the target document before merge</summary>
    public int TargetImageCount { get; init; }

    /// <summary>Images in the source document</summary>
    public int SourceImageCount { get; init; }

    /// <summary>Hyperlinks in the target document before merge</summary>
    public int TargetHyperlinkCount { get; init; }

    /// <summary>Hyperlinks in the source document</summary>
    public int SourceHyperlinkCount { get; init; }
}
