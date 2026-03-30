using System.Text.RegularExpressions;
using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Package;

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
        // Merge resources from source into target
        var imageIdMapping = MergeImages(target, source);
        var hyperlinkIdMapping = MergeHyperlinks(target, source);

        // Optionally add a page break separator
        if (addPageBreak)
        {
            var pageBreakNode = CreatePageBreakNode();
            target.Root.AddChild(pageBreakNode);
        }

        // Clone and append all top-level nodes from source
        foreach (var child in source.Root.Children)
        {
            var clonedNode = CloneNode(child, imageIdMapping, hyperlinkIdMapping);
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
    /// Creates a shallow clone of a document (clones the tree structure but shares PackageData).
    /// For a full independent clone, use CloneDocumentDeep.
    /// </summary>
    /// <param name="source">The document to clone</param>
    /// <returns>A new document with cloned content tree</returns>
    public static WordDocument CloneDocument(WordDocument source)
    {
        var emptyMapping = new Dictionary<string, string>();
        var clonedRoot = CloneNode(source.Root, emptyMapping, emptyMapping);

        // Create new document with cloned root and copy of package data
        var result = new WordDocument(clonedRoot)
        {
            FileName = source.FileName,
            PackageData = ClonePackageData(source.PackageData)
        };

        return result;
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
    public static List<DocumentNode> ExtractSection(this DocumentNode root, string headingText, bool includeNestedHeadings = true)
    {
        var result = new List<DocumentNode>();
        var allNodes = root.Children.ToList();

        // Find the starting heading
        int startIndex = -1;
        int headingLevel = 0;

        for (int i = 0; i < allNodes.Count; i++)
        {
            var node = allNodes[i];
            if (node.Type == ContentType.Heading &&
                node.GetText().Contains(headingText, StringComparison.OrdinalIgnoreCase))
            {
                startIndex = i;
                headingLevel = node.HeadingLevel;
                break;
            }
        }

        if (startIndex < 0)
            return result;

        // Collect nodes until we hit another heading of same or higher level
        for (int i = startIndex; i < allNodes.Count; i++)
        {
            var node = allNodes[i];

            // Check if this is a heading that ends our section
            if (i > startIndex && node.Type == ContentType.Heading)
            {
                if (!includeNestedHeadings)
                {
                    // Stop at any heading
                    break;
                }
                else if (node.HeadingLevel <= headingLevel)
                {
                    // Stop at same level or higher (lower number = higher level)
                    break;
                }
            }

            result.Add(node);
        }

        return result;
    }

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
        // Merge resources from source into target
        var imageIdMapping = MergeImages(target, source);
        var hyperlinkIdMapping = MergeHyperlinks(target, source);

        // Clone and insert nodes
        var nodesToInsert = sourceNodes.ToList();
        var insertIndex = Math.Min(index, parent.Children.Count);

        for (int i = 0; i < nodesToInsert.Count; i++)
        {
            var clonedNode = CloneNode(nodesToInsert[i], imageIdMapping, hyperlinkIdMapping);
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
    public static WordDocument ReplaceSection(
        this WordDocument target,
        string targetHeadingText,
        WordDocument source,
        string sourceHeadingText,
        bool includeNestedHeadings = true)
    {
        // Find and remove the target section
        var targetSection = target.ExtractSection(targetHeadingText, includeNestedHeadings);
        if (targetSection.Count == 0)
            throw new ArgumentException($"Section '{targetHeadingText}' not found in target document", nameof(targetHeadingText));

        var firstNode = targetSection[0];
        var parent = firstNode.Parent;
        if (parent == null)
            throw new InvalidOperationException("Section has no parent");

        var insertIndex = parent.Children.IndexOf(firstNode);

        // Remove the old section
        foreach (var node in targetSection)
        {
            parent.Children.Remove(node);
        }

        // Extract and insert the source section
        var sourceSection = source.ExtractSection(sourceHeadingText, includeNestedHeadings);
        if (sourceSection.Count == 0)
            throw new ArgumentException($"Section '{sourceHeadingText}' not found in source document", nameof(sourceHeadingText));

        return target.InsertNodesAtIndex(parent, insertIndex, sourceSection, source);
    }

    #endregion

    #region Resource Merging

    /// <summary>
    /// Merges images from the source document into the target document.
    /// Returns a mapping from old relationship IDs to new relationship IDs.
    /// </summary>
    private static Dictionary<string, string> MergeImages(WordDocument target, WordDocument source)
    {
        var mapping = new Dictionary<string, string>();

        foreach (var kvp in source.PackageData.Images)
        {
            var oldId = kvp.Key;
            var imageData = kvp.Value;

            // Generate a new unique ID that doesn't conflict with existing IDs
            var newId = GenerateUniqueRelationshipId(target.PackageData.Images.Keys, "rId");

            // Add the image to the target document with the new ID
            target.PackageData.Images[newId] = new ImagePartData
            {
                ContentType = imageData.ContentType,
                Data = imageData.Data,
                OriginalRelationshipId = newId,
                OriginalUri = imageData.OriginalUri
            };

            mapping[oldId] = newId;
        }

        return mapping;
    }

    /// <summary>
    /// Merges hyperlinks from the source document into the target document.
    /// Returns a mapping from old relationship IDs to new relationship IDs.
    /// </summary>
    private static Dictionary<string, string> MergeHyperlinks(WordDocument target, WordDocument source)
    {
        var mapping = new Dictionary<string, string>();

        foreach (var kvp in source.PackageData.HyperlinkRelationships)
        {
            var oldId = kvp.Key;
            var hyperlinkData = kvp.Value;

            // Generate a new unique ID
            var existingIds = target.PackageData.HyperlinkRelationships.Keys
                .Concat(target.PackageData.Images.Keys)
                .ToHashSet();
            var newId = GenerateUniqueRelationshipId(existingIds, "rId");

            target.PackageData.HyperlinkRelationships[newId] = new HyperlinkRelationshipData
            {
                Url = hyperlinkData.Url,
                IsExternal = hyperlinkData.IsExternal
            };

            mapping[oldId] = newId;
        }

        return mapping;
    }

    /// <summary>
    /// Generates a unique relationship ID that doesn't conflict with existing IDs.
    /// </summary>
    private static string GenerateUniqueRelationshipId(IEnumerable<string> existingIds, string prefix)
    {
        var existing = existingIds.ToHashSet();
        var maxId = 0;

        foreach (var id in existing)
        {
            if (id.StartsWith(prefix) && int.TryParse(id[prefix.Length..], out var num))
            {
                maxId = Math.Max(maxId, num);
            }
        }

        // Start from a high number to avoid conflicts
        var newNum = Math.Max(maxId + 1, 1000);
        while (existing.Contains($"{prefix}{newNum}"))
        {
            newNum++;
        }

        return $"{prefix}{newNum}";
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
        var imageIdMapping = MergeImages(target, source);
        var hyperlinkIdMapping = MergeHyperlinks(target, source);

        return nodes.Select(n => CloneNode(n, imageIdMapping, hyperlinkIdMapping)).ToList();
    }

    /// <summary>
    /// Deep clones a single node from a source document, updating resource references for use in the target document.
    /// </summary>
    public static DocumentNode CloneNodeForDocument(
        this WordDocument target,
        WordDocument source,
        DocumentNode node)
    {
        var imageIdMapping = MergeImages(target, source);
        var hyperlinkIdMapping = MergeHyperlinks(target, source);

        return CloneNode(node, imageIdMapping, hyperlinkIdMapping);
    }

    /// <summary>
    /// Deep clones a document node and all its children.
    /// Updates any relationship IDs (images, hyperlinks) using the provided mappings.
    /// </summary>
    internal static DocumentNode CloneNode(
        DocumentNode source,
        Dictionary<string, string> imageIdMapping,
        Dictionary<string, string> hyperlinkIdMapping)
    {
        var clone = new DocumentNode(source.Type)
        {
            Id = Guid.NewGuid().ToString(),
            HeadingLevel = source.HeadingLevel,
            Text = source.Text,
            ParagraphFormatting = source.ParagraphFormatting?.Clone(),
            OriginalXml = UpdateRelationshipIds(source.OriginalXml, imageIdMapping, hyperlinkIdMapping),
            ContentControlProperties = source.ContentControlProperties?.Clone()
        };

        // Clone metadata
        foreach (var kvp in source.Metadata)
        {
            clone.Metadata[kvp.Key] = kvp.Value;
        }

        // Clone formatted runs
        foreach (var run in source.Runs)
        {
            var clonedRun = new FormattedRun(run.Text)
            {
                Formatting = run.Formatting?.Clone() ?? new RunFormatting(),
                IsTab = run.IsTab,
                IsBreak = run.IsBreak,
                BreakType = run.BreakType,
                DocumentPropertyField = run.DocumentPropertyField,
                ContentControlProperties = run.ContentControlProperties?.Clone()
            };
            clone.Runs.Add(clonedRun);
        }

        // Recursively clone children
        // IMPORTANT: Skip Image children if the parent has OriginalXml, because:
        // 1. The OriginalXml already contains the image XML with updated relationship IDs
        // 2. The writer would otherwise write the image twice (once from OriginalXml, once from child node)
        // This prevents duplicate images appearing on top of each other
        var hasOriginalXml = !string.IsNullOrEmpty(source.OriginalXml);

        foreach (var child in source.Children)
        {
            // Skip Image children when parent has OriginalXml - they're already embedded in the XML
            if (hasOriginalXml && child.Type == ContentType.Image)
            {
                continue;
            }

            var clonedChild = CloneNode(child, imageIdMapping, hyperlinkIdMapping);
            clone.AddChild(clonedChild);
        }

        return clone;
    }

    /// <summary>
    /// Updates relationship IDs in OriginalXml to use the new mapped IDs.
    /// </summary>
    private static string? UpdateRelationshipIds(
        string? originalXml,
        Dictionary<string, string> imageIdMapping,
        Dictionary<string, string> hyperlinkIdMapping)
    {
        if (string.IsNullOrEmpty(originalXml))
            return originalXml;

        var result = originalXml;

        // Update image relationship IDs (r:embed="rIdX" or r:link="rIdX")
        foreach (var kvp in imageIdMapping)
        {
            result = Regex.Replace(
                result,
                $@"(r:embed=""){Regex.Escape(kvp.Key)}("")",
                $"$1{kvp.Value}$2");
            result = Regex.Replace(
                result,
                $@"(r:link=""){Regex.Escape(kvp.Key)}("")",
                $"$1{kvp.Value}$2");
        }

        // Update hyperlink relationship IDs (r:id="rIdX")
        foreach (var kvp in hyperlinkIdMapping)
        {
            result = Regex.Replace(
                result,
                $@"(r:id=""){Regex.Escape(kvp.Key)}("")",
                $"$1{kvp.Value}$2");
        }

        return result;
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
            OriginalDocumentXml = source.OriginalDocumentXml
        };

        // Clone core and extended properties
        if (source.CoreProperties != null)
        {
            clone.CoreProperties = new CoreProperties
            {
                Title = source.CoreProperties.Title,
                Subject = source.CoreProperties.Subject,
                Creator = source.CoreProperties.Creator,
                Keywords = source.CoreProperties.Keywords,
                Description = source.CoreProperties.Description,
                LastModifiedBy = source.CoreProperties.LastModifiedBy,
                Revision = source.CoreProperties.Revision,
                Created = source.CoreProperties.Created,
                Modified = source.CoreProperties.Modified,
                Category = source.CoreProperties.Category,
                ContentStatus = source.CoreProperties.ContentStatus
            };
        }

        if (source.ExtendedProperties != null)
        {
            clone.ExtendedProperties = new ExtendedProperties
            {
                Template = source.ExtendedProperties.Template,
                Application = source.ExtendedProperties.Application,
                AppVersion = source.ExtendedProperties.AppVersion,
                Company = source.ExtendedProperties.Company,
                Manager = source.ExtendedProperties.Manager,
                Pages = source.ExtendedProperties.Pages,
                Words = source.ExtendedProperties.Words,
                Characters = source.ExtendedProperties.Characters,
                CharactersWithSpaces = source.ExtendedProperties.CharactersWithSpaces,
                Lines = source.ExtendedProperties.Lines,
                Paragraphs = source.ExtendedProperties.Paragraphs,
                TotalTime = source.ExtendedProperties.TotalTime
            };
        }

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
