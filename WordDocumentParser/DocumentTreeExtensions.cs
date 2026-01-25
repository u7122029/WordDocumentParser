using WordDocumentParser.FormattingModels;

namespace WordDocumentParser;

/// <summary>
/// Extension methods for working with the document tree
/// </summary>
public static class DocumentTreeExtensions
{
    /// <summary>
    /// Finds all nodes matching a predicate
    /// </summary>
    public static IEnumerable<DocumentNode> FindAll(this DocumentNode root, Func<DocumentNode, bool> predicate)
    {
        if (predicate(root))
            yield return root;

        foreach (var child in root.Children)
        {
            foreach (var match in child.FindAll(predicate))
            {
                yield return match;
            }
        }
    }

    /// <summary>
    /// Finds the first node matching a predicate
    /// </summary>
    public static DocumentNode? FindFirst(this DocumentNode root, Func<DocumentNode, bool> predicate)
        => root.FindAll(predicate).FirstOrDefault();

    /// <summary>
    /// Gets all headings at a specific level
    /// </summary>
    public static IEnumerable<DocumentNode> GetHeadingsAtLevel(this DocumentNode root, int level)
        => root.FindAll(n => n.Type == ContentType.Heading && n.HeadingLevel == level);

    /// <summary>
    /// Gets all headings in the document
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllHeadings(this DocumentNode root)
        => root.FindAll(n => n.Type == ContentType.Heading);

    /// <summary>
    /// Gets all tables in the document
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllTables(this DocumentNode root)
        => root.FindAll(n => n.Type == ContentType.Table);

    /// <summary>
    /// Gets all images in the document
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllImages(this DocumentNode root)
        => root.FindAll(n => n.Type == ContentType.Image);

    /// <summary>
    /// Gets the table of contents as a flat list with indentation info
    /// </summary>
    public static List<(int Level, string Title, DocumentNode Node)> GetTableOfContents(this DocumentNode root)
        => [.. root.GetAllHeadings().Select(h => (h.HeadingLevel, h.Text, h))];

    /// <summary>
    /// Gets all text content under a node (recursive)
    /// </summary>
    public static string GetAllText(this DocumentNode node)
    {
        var texts = new List<string>();

        if (!string.IsNullOrEmpty(node.Text) && node.Type is not ContentType.Table and not ContentType.Image)
        {
            texts.Add(node.Text);
        }

        foreach (var child in node.Children)
        {
            texts.Add(child.GetAllText());
        }

        return string.Join("\n", texts.Where(t => !string.IsNullOrWhiteSpace(t)));
    }

    /// <summary>
    /// Gets the path from root to this node
    /// </summary>
    public static List<DocumentNode> GetPath(this DocumentNode node)
    {
        var path = new List<DocumentNode>();
        var current = node;

        while (current is not null)
        {
            path.Insert(0, current);
            current = current.Parent;
        }

        return path;
    }

    /// <summary>
    /// Gets the heading path (breadcrumb) for a node
    /// </summary>
    public static string GetHeadingPath(this DocumentNode node, string separator = " > ")
    {
        var headings = node.GetPath()
            .Where(n => n.Type is ContentType.Heading or ContentType.Document)
            .Select(n => n.Text)
            .Where(t => !string.IsNullOrEmpty(t));

        return string.Join(separator, headings);
    }

    /// <summary>
    /// Gets siblings of a node
    /// </summary>
    public static IEnumerable<DocumentNode> GetSiblings(this DocumentNode node)
        => node.Parent?.Children.Where(c => c != node) ?? [];

    /// <summary>
    /// Gets the next sibling
    /// </summary>
    public static DocumentNode? GetNextSibling(this DocumentNode node)
    {
        if (node.Parent is null) return null;

        var siblings = node.Parent.Children;
        var index = siblings.IndexOf(node);

        return index >= 0 && index < siblings.Count - 1 ? siblings[index + 1] : null;
    }

    /// <summary>
    /// Gets the previous sibling
    /// </summary>
    public static DocumentNode? GetPreviousSibling(this DocumentNode node)
    {
        if (node.Parent is null) return null;

        var siblings = node.Parent.Children;
        var index = siblings.IndexOf(node);

        return index > 0 ? siblings[index - 1] : null;
    }

    /// <summary>
    /// Counts nodes by type
    /// </summary>
    public static Dictionary<ContentType, int> CountByType(this DocumentNode root)
    {
        var counts = new Dictionary<ContentType, int>();

        foreach (var node in root.FindAll(_ => true))
        {
            counts.TryGetValue(node.Type, out var count);
            counts[node.Type] = count + 1;
        }

        return counts;
    }

    /// <summary>
    /// Flattens the tree to a list in document order
    /// </summary>
    public static List<DocumentNode> Flatten(this DocumentNode root)
    {
        List<DocumentNode> result = [root];

        foreach (var child in root.Children)
        {
            result.AddRange(child.Flatten());
        }

        return result;
    }

    /// <summary>
    /// Gets a section by heading text (case-insensitive partial match)
    /// </summary>
    public static DocumentNode? GetSection(this DocumentNode root, string headingText)
        => root.FindFirst(n =>
            n.Type == ContentType.Heading &&
            n.Text.Contains(headingText, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Extracts TableData from a table node
    /// </summary>
    public static TableData? GetTableData(this DocumentNode tableNode)
    {
        if (tableNode.Type != ContentType.Table)
            return null;

        return tableNode.Metadata.TryGetValue("TableData", out var data)
            ? data as TableData
            : null;
    }

    /// <summary>
    /// Extracts ImageData from an image node
    /// </summary>
    public static ImageData? GetImageData(this DocumentNode imageNode)
    {
        if (imageNode.Type != ContentType.Image)
            return null;

        return imageNode.Metadata.TryGetValue("ImageData", out var data)
            ? data as ImageData
            : null;
    }

    /// <summary>
    /// Gets all content controls in the document (both block-level and inline)
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllContentControls(this DocumentNode root)
        => root.FindAll(n => n.IsContentControl || n.HasInlineContentControls());

    /// <summary>
    /// Gets all content controls of a specific type (both block-level and inline)
    /// </summary>
    public static IEnumerable<DocumentNode> GetContentControlsByType(this DocumentNode root, ContentControlType type)
        => root.FindAll(n => n.ContentControlProperties?.Type == type ||
                             n.Runs.Any(r => r.ContentControlProperties?.Type == type));

    /// <summary>
    /// Checks if a node has inline content controls in its runs
    /// </summary>
    public static bool HasInlineContentControls(this DocumentNode node)
        => node.Runs.Any(r => r.IsContentControlRun);

    /// <summary>
    /// Gets all inline content control properties from a node's runs
    /// </summary>
    public static IEnumerable<ContentControlProperties> GetInlineContentControlProperties(this DocumentNode node)
        => node.Runs
            .Where(r => r.IsContentControlRun && r.ContentControlProperties is not null)
            .Select(r => r.ContentControlProperties!)
            .Distinct();

    /// <summary>
    /// Finds a content control by its tag (checks both block-level and inline controls)
    /// </summary>
    public static DocumentNode? FindContentControlByTag(this DocumentNode root, string tag)
        => root.FindFirst(n =>
            n.ContentControlProperties?.Tag == tag ||
            n.Runs.Any(r => r.ContentControlProperties?.Tag == tag));

    /// <summary>
    /// Finds a content control by its alias/title (checks both block-level and inline controls)
    /// </summary>
    public static DocumentNode? FindContentControlByAlias(this DocumentNode root, string alias)
        => root.FindFirst(n =>
            n.ContentControlProperties?.Alias == alias ||
            n.Runs.Any(r => r.ContentControlProperties?.Alias == alias));

    /// <summary>
    /// Finds a content control by its ID (checks both block-level and inline controls)
    /// </summary>
    public static DocumentNode? FindContentControlById(this DocumentNode root, int id)
        => root.FindFirst(n =>
            n.ContentControlProperties?.Id == id ||
            n.Runs.Any(r => r.ContentControlProperties?.Id == id));

    /// <summary>
    /// Gets all document property content controls
    /// </summary>
    public static IEnumerable<DocumentNode> GetDocumentPropertyControls(this DocumentNode root)
        => root.GetContentControlsByType(ContentControlType.DocumentProperty);

    /// <summary>
    /// Gets all nodes that contain document property fields
    /// </summary>
    public static IEnumerable<DocumentNode> GetNodesWithDocumentPropertyFields(this DocumentNode root)
        => root.FindAll(n => n.HasDocumentPropertyFields);

    /// <summary>
    /// Gets all document property fields in the document
    /// </summary>
    public static IEnumerable<DocumentPropertyField> GetAllDocumentPropertyFields(this DocumentNode root)
    {
        foreach (var node in root.FindAll(_ => true))
        {
            foreach (var run in node.Runs)
            {
                if (run.DocumentPropertyField is not null)
                {
                    yield return run.DocumentPropertyField;
                }
            }
        }
    }

    /// <summary>
    /// Gets all text from the document with metadata annotations for document properties and content controls
    /// </summary>
    public static string GetAllTextWithMetadata(this DocumentNode node)
    {
        var texts = new List<string>();

        if (node.Type is not ContentType.Table and not ContentType.Image)
        {
            var text = node.GetTextWithMetadata();
            if (!string.IsNullOrEmpty(text))
            {
                texts.Add(text);
            }
        }

        foreach (var child in node.Children)
        {
            texts.Add(child.GetAllTextWithMetadata());
        }

        return string.Join("\n", texts.Where(t => !string.IsNullOrWhiteSpace(t)));
    }

    /// <summary>
    /// Sets the value of a content control by tag
    /// </summary>
    public static bool SetContentControlValueByTag(this DocumentNode root, string tag, string newValue)
    {
        var control = root.FindContentControlByTag(tag);
        if (control is null) return false;

        control.Text = newValue;
        if (control.ContentControlProperties is not null)
        {
            control.ContentControlProperties.Value = newValue;
        }

        // Update runs if present
        if (control.Runs.Count > 0)
        {
            control.Runs.Clear();
            control.Runs.Add(new FormattedRun(newValue));
        }

        return true;
    }

    /// <summary>
    /// Sets the value of a content control by alias
    /// </summary>
    public static bool SetContentControlValueByAlias(this DocumentNode root, string alias, string newValue)
    {
        var control = root.FindContentControlByAlias(alias);
        if (control is null) return false;

        control.Text = newValue;
        if (control.ContentControlProperties is not null)
        {
            control.ContentControlProperties.Value = newValue;
        }

        // Update runs if present
        if (control.Runs.Count > 0)
        {
            control.Runs.Clear();
            control.Runs.Add(new FormattedRun(newValue));
        }

        return true;
    }

    /// <summary>
    /// Gets all content control tags in the document
    /// </summary>
    public static IEnumerable<string> GetContentControlTags(this DocumentNode root)
        => root.GetAllContentControls()
            .Where(n => !string.IsNullOrEmpty(n.ContentControlProperties?.Tag))
            .Select(n => n.ContentControlProperties!.Tag!);

    /// <summary>
    /// Gets content control properties by tag
    /// </summary>
    public static ContentControlProperties? GetContentControlPropertiesByTag(this DocumentNode root, string tag)
        => root.FindContentControlByTag(tag)?.ContentControlProperties;

    /// <summary>
    /// Saves the document tree to a Word document file (.docx)
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <param name="filePath">The path where the document will be saved</param>
    public static void SaveToFile(this DocumentNode root, string filePath)
    {
        using var writer = new WordDocumentTreeWriter();
        writer.WriteToFile(root, filePath);
    }

    /// <summary>
    /// Saves the document tree to a stream as a Word document (.docx)
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <param name="stream">The stream to write to</param>
    public static void SaveToStream(this DocumentNode root, Stream stream)
    {
        using var writer = new WordDocumentTreeWriter();
        writer.WriteToStream(root, stream);
    }

    /// <summary>
    /// Saves the document tree to a byte array as a Word document (.docx)
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <returns>The document as a byte array</returns>
    public static byte[] ToDocxBytes(this DocumentNode root)
    {
        using var stream = new MemoryStream();
        root.SaveToStream(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Removes a content control from a node, keeping the text content but removing the control wrapper.
    /// For block-level content controls, this clears the ContentControlProperties and OriginalXml.
    /// For inline content controls, this removes the ContentControlProperties from the affected runs.
    /// </summary>
    /// <param name="node">The node containing the content control</param>
    /// <param name="contentControlId">The ID of the content control to remove. If null, removes all content controls from the node.</param>
    /// <returns>True if any content control was removed, false otherwise</returns>
    public static bool RemoveContentControl(this DocumentNode node, int? contentControlId = null)
    {
        var removed = false;

        // Handle block-level content control
        if (node.ContentControlProperties is not null)
        {
            if (contentControlId is null || node.ContentControlProperties.Id == contentControlId)
            {
                node.ContentControlProperties = null;
                node.OriginalXml = null; // Force rebuild without SDT wrapper
                node.Metadata.Remove("IsSdtContent");
                node.Metadata.Remove("IsSdtBlock");
                removed = true;
            }
        }

        // Handle inline content controls in runs
        foreach (var run in node.Runs)
        {
            if (run.ContentControlProperties is not null)
            {
                if (contentControlId is null || run.ContentControlProperties.Id == contentControlId)
                {
                    run.ContentControlProperties = null;
                    removed = true;
                }
            }
        }

        // If we removed any inline content controls, we need to clear OriginalXml to force rebuild
        if (removed && node.Runs.Any())
        {
            node.OriginalXml = null;
        }

        return removed;
    }

    /// <summary>
    /// Removes all content controls from a document, keeping the text content.
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <returns>The number of content controls removed</returns>
    public static int RemoveAllContentControls(this DocumentNode root)
    {
        var count = 0;
        foreach (var node in root.FindAll(_ => true))
        {
            if (node.RemoveContentControl())
            {
                count++;
            }
        }
        return count;
    }

    /// <summary>
    /// Removes a document property field from a node's runs, keeping the text content.
    /// </summary>
    /// <param name="node">The node containing the document property field</param>
    /// <param name="propertyName">The name of the property to remove. If null, removes all document property fields.</param>
    /// <returns>True if any document property field was removed, false otherwise</returns>
    public static bool RemoveDocumentPropertyField(this DocumentNode node, string? propertyName = null)
    {
        var removed = false;

        foreach (var run in node.Runs)
        {
            if (run.DocumentPropertyField is not null)
            {
                if (propertyName is null ||
                    string.Equals(run.DocumentPropertyField.PropertyName, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    run.DocumentPropertyField = null;
                    removed = true;
                }
            }
        }

        // If we removed any fields, clear OriginalXml to force rebuild
        if (removed)
        {
            node.OriginalXml = null;
        }

        return removed;
    }

    /// <summary>
    /// Removes all document property fields from a document, keeping the text content.
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <returns>The number of nodes that had document property fields removed</returns>
    public static int RemoveAllDocumentPropertyFields(this DocumentNode root)
    {
        var count = 0;
        foreach (var node in root.FindAll(_ => true))
        {
            if (node.RemoveDocumentPropertyField())
            {
                count++;
            }
        }
        return count;
    }

    /// <summary>
    /// Removes a content control by its tag, keeping the text content.
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <param name="tag">The tag of the content control to remove</param>
    /// <returns>True if the content control was found and removed, false otherwise</returns>
    public static bool RemoveContentControlByTag(this DocumentNode root, string tag)
    {
        var node = root.FindContentControlByTag(tag);
        if (node is null) return false;

        // Check block-level first, then inline controls
        var ccId = node.ContentControlProperties?.Tag == tag
            ? node.ContentControlProperties?.Id
            : node.Runs.FirstOrDefault(r => r.ContentControlProperties?.Tag == tag)?.ContentControlProperties?.Id;

        return node.RemoveContentControl(ccId);
    }

    /// <summary>
    /// Removes a content control by its alias, keeping the text content.
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <param name="alias">The alias of the content control to remove</param>
    /// <returns>True if the content control was found and removed, false otherwise</returns>
    public static bool RemoveContentControlByAlias(this DocumentNode root, string alias)
    {
        var node = root.FindContentControlByAlias(alias);
        if (node is null) return false;

        // Check block-level first, then inline controls
        var ccId = node.ContentControlProperties?.Alias == alias
            ? node.ContentControlProperties?.Id
            : node.Runs.FirstOrDefault(r => r.ContentControlProperties?.Alias == alias)?.ContentControlProperties?.Id;

        return node.RemoveContentControl(ccId);
    }
}
