using WordDocumentParser.Core;
using WordDocumentParser.Models.Images;
using WordDocumentParser.Models.Tables;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for querying document content by type.
/// </summary>
public static class TreeQueryExtensions
{
    /// <summary>
    /// Gets all headings at a specific level (H1, H2, etc.).
    /// </summary>
    /// <param name="root">Root node to search from</param>
    /// <param name="level">Heading level (1-9)</param>
    public static IEnumerable<DocumentNode> GetHeadingsAtLevel(this DocumentNode root, int level)
        => root.FindAll(n => n.Type == ContentType.Heading && n.HeadingLevel == level);

    /// <summary>
    /// Gets all headings in the document.
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllHeadings(this DocumentNode root)
        => root.FindAll(n => n.Type == ContentType.Heading);

    /// <summary>
    /// Gets all tables in the document.
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllTables(this DocumentNode root)
        => root.FindAll(n => n.Type == ContentType.Table);

    /// <summary>
    /// Gets all images in the document.
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllImages(this DocumentNode root)
        => root.FindAll(n => n.Type == ContentType.Image);

    /// <summary>
    /// Gets the table of contents as a flat list with level info.
    /// </summary>
    /// <returns>List of tuples (HeadingLevel, HeadingText, Node)</returns>
    public static List<(int Level, string Title, DocumentNode Node)> GetTableOfContents(this DocumentNode root)
        => [.. root.GetAllHeadings().Select(h => (h.HeadingLevel, h.Text, h))];

    /// <summary>
    /// Gets all text content under a node (recursive).
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
    /// Counts nodes by content type.
    /// </summary>
    /// <returns>Dictionary mapping ContentType to count</returns>
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
    /// Extracts TableData from a table node.
    /// </summary>
    /// <returns>TableData if the node is a table, null otherwise</returns>
    public static TableData? GetTableData(this DocumentNode tableNode)
    {
        if (tableNode.Type != ContentType.Table)
            return null;

        return tableNode.Metadata.TryGetValue("TableData", out var data)
            ? data as TableData
            : null;
    }

    /// <summary>
    /// Extracts ImageData from an image node.
    /// </summary>
    /// <returns>ImageData if the node is an image, null otherwise</returns>
    public static ImageData? GetImageData(this DocumentNode imageNode)
    {
        if (imageNode.Type != ContentType.Image)
            return null;

        return imageNode.Metadata.TryGetValue("ImageData", out var data)
            ? data as ImageData
            : null;
    }
}
