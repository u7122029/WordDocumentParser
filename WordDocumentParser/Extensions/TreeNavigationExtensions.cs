using WordDocumentParser.Core;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for navigating the document tree structure.
/// </summary>
public static class TreeNavigationExtensions
{
    /// <summary>
    /// Finds all nodes matching a predicate (depth-first traversal).
    /// </summary>
    /// <param name="root">Starting node for the search</param>
    /// <param name="predicate">Condition to match</param>
    /// <returns>All matching nodes in document order</returns>
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
    /// Finds the first node matching a predicate.
    /// </summary>
    public static DocumentNode? FindFirst(this DocumentNode root, Func<DocumentNode, bool> predicate)
        => root.FindAll(predicate).FirstOrDefault();

    /// <summary>
    /// Gets the path from root to this node.
    /// </summary>
    /// <returns>List of nodes from root to current node</returns>
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
    /// Gets the heading path (breadcrumb) for a node.
    /// </summary>
    /// <param name="node">The node to get the path for</param>
    /// <param name="separator">Separator between path segments</param>
    /// <returns>String like "Document > Chapter 1 > Section 1.1"</returns>
    public static string GetHeadingPath(this DocumentNode node, string separator = " > ")
    {
        var headings = node.GetPath()
            .Where(n => n.Type is ContentType.Heading or ContentType.Document)
            .Select(n => n.Text)
            .Where(t => !string.IsNullOrEmpty(t));

        return string.Join(separator, headings);
    }

    /// <summary>
    /// Gets siblings of a node (excluding the node itself).
    /// </summary>
    public static IEnumerable<DocumentNode> GetSiblings(this DocumentNode node)
        => node.Parent?.Children.Where(c => c != node) ?? [];

    /// <summary>
    /// Gets the next sibling in document order.
    /// </summary>
    public static DocumentNode? GetNextSibling(this DocumentNode node)
    {
        if (node.Parent is null) return null;

        var siblings = node.Parent.Children;
        var index = siblings.IndexOf(node);

        return index >= 0 && index < siblings.Count - 1 ? siblings[index + 1] : null;
    }

    /// <summary>
    /// Gets the previous sibling in document order.
    /// </summary>
    public static DocumentNode? GetPreviousSibling(this DocumentNode node)
    {
        if (node.Parent is null) return null;

        var siblings = node.Parent.Children;
        var index = siblings.IndexOf(node);

        return index > 0 ? siblings[index - 1] : null;
    }

    /// <summary>
    /// Flattens the tree to a list in document order.
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
    /// Gets a section by heading text (case-insensitive partial match).
    /// </summary>
    public static DocumentNode? GetSection(this DocumentNode root, string headingText)
        => root.FindFirst(n =>
            n.Type == ContentType.Heading &&
            n.Text.Contains(headingText, StringComparison.OrdinalIgnoreCase));
}
