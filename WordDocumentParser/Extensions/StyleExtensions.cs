using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for querying and modifying paragraph styles.
/// </summary>
public static class StyleExtensions
{
    #region Finding nodes by style

    /// <summary>
    /// Finds all nodes with a specific paragraph style.
    /// </summary>
    /// <param name="document">The document to search</param>
    /// <param name="styleId">The style ID to search for (e.g., "Heading1", "Normal", "Quote")</param>
    /// <returns>All nodes matching the specified style</returns>
    public static IEnumerable<DocumentNode> FindByStyle(this WordDocument document, string styleId)
        => document.Root.FindByStyle(styleId);

    /// <summary>
    /// Finds all nodes with a specific paragraph style.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="styleId">The style ID to search for (e.g., "Heading1", "Normal", "Quote")</param>
    /// <returns>All nodes matching the specified style</returns>
    public static IEnumerable<DocumentNode> FindByStyle(this DocumentNode root, string styleId)
    {
        return root.FindAll(n =>
            n.ParagraphFormatting?.StyleId != null &&
            n.ParagraphFormatting.StyleId.Equals(styleId, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Finds all nodes that have any of the specified styles.
    /// </summary>
    /// <param name="document">The document to search</param>
    /// <param name="styleIds">The style IDs to search for</param>
    /// <returns>All nodes matching any of the specified styles</returns>
    public static IEnumerable<DocumentNode> FindByStyles(this WordDocument document, params string[] styleIds)
        => document.Root.FindByStyles(styleIds);

    /// <summary>
    /// Finds all nodes that have any of the specified styles.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="styleIds">The style IDs to search for</param>
    /// <returns>All nodes matching any of the specified styles</returns>
    public static IEnumerable<DocumentNode> FindByStyles(this DocumentNode root, params string[] styleIds)
    {
        var styleSet = new HashSet<string>(styleIds, StringComparer.OrdinalIgnoreCase);
        return root.FindAll(n =>
            n.ParagraphFormatting?.StyleId != null &&
            styleSet.Contains(n.ParagraphFormatting.StyleId));
    }

    #endregion

    #region Changing styles

    /// <summary>
    /// Changes the paragraph style of a node.
    /// </summary>
    /// <param name="node">The node to modify</param>
    /// <param name="newStyleId">The new style ID (e.g., "Heading2", "Quote", "NoSpacing")</param>
    /// <remarks>
    /// <para>
    /// The change is recorded on the node's formatting; the writer inserts <c>w:pStyle</c> into the
    /// <c>w:pPr</c> of the node's original XML on save, at its schema position.
    /// </para>
    /// <para>
    /// This changes the node's type and heading level to match the new style, but does not move it
    /// in the tree — the node keeps its current parent and children.
    /// </para>
    /// </remarks>
    public static void ChangeStyle(this DocumentNode node, string newStyleId)
    {
        node.ParagraphFormatting ??= new ParagraphFormatting();
        node.ParagraphFormatting.StyleId = newStyleId;

        // Assigning the style a node already had is still an instruction to write it.
        node.ParagraphFormatting.MarkChanged(nameof(ParagraphFormatting.StyleId));

        if (newStyleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(newStyleId.AsSpan(7), out var level))
        {
            node.HeadingLevel = level;
            node.Type = ContentType.Heading;
        }
        else if (node.Type == ContentType.Heading &&
                 !newStyleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
        {
            // Changing from a heading to a non-heading style
            node.HeadingLevel = 0;
            node.Type = ContentType.Paragraph;
        }
    }

    /// <summary>
    /// Changes the style of all nodes matching a specific style.
    /// </summary>
    /// <param name="document">The document to modify</param>
    /// <param name="fromStyleId">The style to search for</param>
    /// <param name="toStyleId">The style to change to</param>
    /// <returns>The number of nodes changed</returns>
    public static int ChangeStyleBulk(this WordDocument document, string fromStyleId, string toStyleId)
        => document.Root.ChangeStyleBulk(fromStyleId, toStyleId);

    /// <summary>
    /// Changes the style of all nodes matching a specific style.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="fromStyleId">The style to search for</param>
    /// <param name="toStyleId">The style to change to</param>
    /// <returns>The number of nodes changed</returns>
    public static int ChangeStyleBulk(this DocumentNode root, string fromStyleId, string toStyleId)
    {
        var nodes = root.FindByStyle(fromStyleId).ToList();
        foreach (var node in nodes)
        {
            node.ChangeStyle(toStyleId);
        }
        return nodes.Count;
    }

    /// <summary>
    /// Changes the style of all nodes matching a predicate.
    /// </summary>
    /// <param name="document">The document to modify</param>
    /// <param name="predicate">The condition to match nodes</param>
    /// <param name="toStyleId">The style to change to</param>
    /// <returns>The number of nodes changed</returns>
    public static int ChangeStyleWhere(this WordDocument document, Func<DocumentNode, bool> predicate, string toStyleId)
        => document.Root.ChangeStyleWhere(predicate, toStyleId);

    /// <summary>
    /// Changes the style of all nodes matching a predicate.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="predicate">The condition to match nodes</param>
    /// <param name="toStyleId">The style to change to</param>
    /// <returns>The number of nodes changed</returns>
    public static int ChangeStyleWhere(this DocumentNode root, Func<DocumentNode, bool> predicate, string toStyleId)
    {
        var nodes = root.FindAll(predicate).ToList();
        foreach (var node in nodes)
        {
            node.ChangeStyle(toStyleId);
        }
        return nodes.Count;
    }

    #endregion

    #region Style statistics

    /// <summary>
    /// Gets a dictionary of style IDs to their occurrence counts.
    /// </summary>
    /// <param name="document">The document to analyze</param>
    /// <returns>Dictionary mapping style names to counts</returns>
    public static Dictionary<string, int> GetStyleDistribution(this WordDocument document)
        => document.Root.GetStyleDistribution();

    /// <summary>
    /// Gets a dictionary of style IDs to their occurrence counts.
    /// </summary>
    /// <param name="root">The root node to analyze</param>
    /// <returns>Dictionary mapping style names to counts</returns>
    public static Dictionary<string, int> GetStyleDistribution(this DocumentNode root)
    {
        var distribution = new Dictionary<string, int>();

        foreach (var node in root.FindAllContent(_ => true))
        {
            if (node.Type is ContentType.Paragraph or ContentType.Heading or ContentType.ListItem)
            {
                var styleId = node.ParagraphFormatting?.StyleId ?? "(no style)";
                distribution[styleId] = distribution.GetValueOrDefault(styleId, 0) + 1;
            }
        }

        return distribution;
    }

    /// <summary>
    /// Gets the style of a node, or null if not set.
    /// </summary>
    /// <param name="node">The node to check</param>
    /// <returns>The style ID or null</returns>
    public static string? GetStyle(this DocumentNode node)
        => node.ParagraphFormatting?.StyleId;

    /// <summary>
    /// Checks if a node has a specific style.
    /// </summary>
    /// <param name="node">The node to check</param>
    /// <param name="styleId">The style ID to check for</param>
    /// <returns>True if the node has the specified style</returns>
    public static bool HasStyle(this DocumentNode node, string styleId)
        => node.ParagraphFormatting?.StyleId?.Equals(styleId, StringComparison.OrdinalIgnoreCase) ?? false;

    /// <summary>
    /// Checks if a node has any of the specified styles.
    /// </summary>
    /// <param name="node">The node to check</param>
    /// <param name="styleIds">The style IDs to check for</param>
    /// <returns>True if the node has any of the specified styles</returns>
    public static bool HasAnyStyle(this DocumentNode node, params string[] styleIds)
    {
        var style = node.ParagraphFormatting?.StyleId;
        if (style == null) return false;
        return styleIds.Any(s => s.Equals(style, StringComparison.OrdinalIgnoreCase));
    }

    #endregion
}
