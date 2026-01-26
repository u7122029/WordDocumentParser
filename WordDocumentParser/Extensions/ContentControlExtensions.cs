using WordDocumentParser.Models.ContentControls;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for working with content controls (SDT - Structured Document Tags).
/// </summary>
public static class ContentControlExtensions
{
    /// <summary>
    /// Gets all content controls in the document (both block-level and inline).
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllContentControls(this DocumentNode root)
        => root.FindAll(n => n.IsContentControl || n.HasInlineContentControls());

    /// <summary>
    /// Gets all content controls of a specific type.
    /// </summary>
    public static IEnumerable<DocumentNode> GetContentControlsByType(this DocumentNode root, ContentControlType type)
        => root.FindAll(n => n.ContentControlProperties?.Type == type ||
                             n.Runs.Any(r => r.ContentControlProperties?.Type == type));

    /// <summary>
    /// Checks if a node has inline content controls in its runs.
    /// </summary>
    public static bool HasInlineContentControls(this DocumentNode node)
        => node.Runs.Any(r => r.IsContentControlRun);

    /// <summary>
    /// Gets all inline content control properties from a node's runs.
    /// </summary>
    public static IEnumerable<ContentControlProperties> GetInlineContentControlProperties(this DocumentNode node)
        => node.Runs
            .Where(r => r.IsContentControlRun && r.ContentControlProperties is not null)
            .Select(r => r.ContentControlProperties!)
            .Distinct();

    /// <summary>
    /// Finds a content control by its tag (checks both block-level and inline).
    /// </summary>
    public static DocumentNode? FindContentControlByTag(this DocumentNode root, string tag)
        => root.FindFirst(n =>
            n.ContentControlProperties?.Tag == tag ||
            n.Runs.Any(r => r.ContentControlProperties?.Tag == tag));

    /// <summary>
    /// Finds a content control by its alias/title (checks both block-level and inline).
    /// </summary>
    public static DocumentNode? FindContentControlByAlias(this DocumentNode root, string alias)
        => root.FindFirst(n =>
            n.ContentControlProperties?.Alias == alias ||
            n.Runs.Any(r => r.ContentControlProperties?.Alias == alias));

    /// <summary>
    /// Finds a content control by its ID (checks both block-level and inline).
    /// </summary>
    public static DocumentNode? FindContentControlById(this DocumentNode root, int id)
        => root.FindFirst(n =>
            n.ContentControlProperties?.Id == id ||
            n.Runs.Any(r => r.ContentControlProperties?.Id == id));

    /// <summary>
    /// Sets the value of a content control by tag.
    /// </summary>
    /// <returns>True if control was found and updated</returns>
    public static bool SetContentControlValueByTag(this DocumentNode root, string tag, string newValue)
    {
        var control = root.FindContentControlByTag(tag);
        if (control is null) return false;

        control.Text = newValue;
        if (control.ContentControlProperties is not null)
        {
            control.ContentControlProperties.Value = newValue;
        }

        if (control.Runs.Count > 0)
        {
            control.Runs.Clear();
            control.Runs.Add(new FormattedRun(newValue));
        }

        return true;
    }

    /// <summary>
    /// Sets the value of a content control by alias.
    /// </summary>
    /// <returns>True if control was found and updated</returns>
    public static bool SetContentControlValueByAlias(this DocumentNode root, string alias, string newValue)
    {
        var control = root.FindContentControlByAlias(alias);
        if (control is null) return false;

        control.Text = newValue;
        if (control.ContentControlProperties is not null)
        {
            control.ContentControlProperties.Value = newValue;
        }

        if (control.Runs.Count > 0)
        {
            control.Runs.Clear();
            control.Runs.Add(new FormattedRun(newValue));
        }

        return true;
    }

    /// <summary>
    /// Gets all content control tags in the document.
    /// </summary>
    public static IEnumerable<string> GetContentControlTags(this DocumentNode root)
        => root.GetAllContentControls()
            .Where(n => !string.IsNullOrEmpty(n.ContentControlProperties?.Tag))
            .Select(n => n.ContentControlProperties!.Tag!);

    /// <summary>
    /// Gets content control properties by tag.
    /// </summary>
    public static ContentControlProperties? GetContentControlPropertiesByTag(this DocumentNode root, string tag)
        => root.FindContentControlByTag(tag)?.ContentControlProperties;

    /// <summary>
    /// Removes a content control from a node, keeping the text content.
    /// </summary>
    /// <param name="node">The node containing the content control</param>
    /// <param name="contentControlId">ID to remove, or null to remove all</param>
    /// <returns>True if any content control was removed</returns>
    public static bool RemoveContentControl(this DocumentNode node, int? contentControlId = null)
    {
        var removed = false;

        // Handle block-level content control
        if (node.ContentControlProperties is not null)
        {
            if (contentControlId is null || node.ContentControlProperties.Id == contentControlId)
            {
                node.ContentControlProperties = null;
                node.OriginalXml = null;
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

        if (removed && node.Runs.Any())
        {
            node.OriginalXml = null;
        }

        return removed;
    }

    /// <summary>
    /// Removes all content controls from a document, keeping text content.
    /// </summary>
    /// <returns>Number of nodes modified</returns>
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
    /// Removes a content control by its tag, keeping text content.
    /// </summary>
    public static bool RemoveContentControlByTag(this DocumentNode root, string tag)
    {
        var node = root.FindContentControlByTag(tag);
        if (node is null) return false;

        var ccId = node.ContentControlProperties?.Tag == tag
            ? node.ContentControlProperties?.Id
            : node.Runs.FirstOrDefault(r => r.ContentControlProperties?.Tag == tag)?.ContentControlProperties?.Id;

        return node.RemoveContentControl(ccId);
    }

    /// <summary>
    /// Removes a content control by its alias, keeping text content.
    /// </summary>
    public static bool RemoveContentControlByAlias(this DocumentNode root, string alias)
    {
        var node = root.FindContentControlByAlias(alias);
        if (node is null) return false;

        var ccId = node.ContentControlProperties?.Alias == alias
            ? node.ContentControlProperties?.Id
            : node.Runs.FirstOrDefault(r => r.ContentControlProperties?.Alias == alias)?.ContentControlProperties?.Id;

        return node.RemoveContentControl(ccId);
    }
}
