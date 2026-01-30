using System.Text.RegularExpressions;
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
    /// Updates both the ParagraphFormatting.StyleId and the OriginalXml if present.
    /// </summary>
    /// <param name="node">The node to modify</param>
    /// <param name="newStyleId">The new style ID (e.g., "Heading2", "Quote", "NoSpacing")</param>
    public static void ChangeStyle(this DocumentNode node, string newStyleId)
    {
        // Ensure ParagraphFormatting exists
        node.ParagraphFormatting ??= new ParagraphFormatting();

        var oldStyleId = node.ParagraphFormatting.StyleId;

        // Update the StyleId
        node.ParagraphFormatting.StyleId = newStyleId;

        // If there's OriginalXml, update the style in the XML as well
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            node.OriginalXml = UpdateStyleInXml(node.OriginalXml, oldStyleId, newStyleId);
        }

        // If the new style is a heading, update the HeadingLevel and Type
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

        foreach (var node in root.FindAll(_ => true))
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

    #region Private helpers

    /// <summary>
    /// Updates the style ID in an XML string.
    /// </summary>
    private static string UpdateStyleInXml(string xml, string? oldStyleId, string newStyleId)
    {
        if (oldStyleId != null)
        {
            // Replace existing style: <w:pStyle w:val="OldStyle"/>
            var pattern = $@"<w:pStyle\s+w:val=""{Regex.Escape(oldStyleId)}""";
            var replacement = $@"<w:pStyle w:val=""{newStyleId}""";
            xml = Regex.Replace(xml, pattern, replacement, RegexOptions.IgnoreCase);
        }
        else
        {
            // No existing style - need to add one
            // Look for <w:pPr> and add the style inside it
            var pPrPattern = @"(<w:pPr[^>]*>)";
            var match = Regex.Match(xml, pPrPattern);
            if (match.Success)
            {
                var pPrTag = match.Groups[1].Value;
                xml = xml.Replace(pPrTag, $@"{pPrTag}<w:pStyle w:val=""{newStyleId}""/>");
            }
            else
            {
                // No pPr exists - need to add one after <w:p ...>
                var pTagPattern = @"(<w:p[^>]*>)";
                match = Regex.Match(xml, pTagPattern);
                if (match.Success)
                {
                    var pTag = match.Groups[1].Value;
                    xml = xml.Replace(pTag, $@"{pTag}<w:pPr><w:pStyle w:val=""{newStyleId}""/></w:pPr>");
                }
            }
        }

        return xml;
    }

    #endregion
}
