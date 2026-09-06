using WordDocumentParser.Core;
using WordDocumentParser.Models.ContentControls;
using WordDocumentParser.Models.Formatting;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for working with content controls (SDT - Structured Document Tags).
/// </summary>
public static class ContentControlExtensions
{
    /// <summary>
    /// Gets all content controls in the document (both block-level and inline).
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllContentControls(this WordDocument document)
        => document.Root.FindAll(n => n.IsContentControl || n.HasInlineContentControls());

    /// <summary>
    /// Gets all content controls in the node tree (both block-level and inline).
    /// </summary>
    public static IEnumerable<DocumentNode> GetAllContentControls(this DocumentNode root)
        => root.FindAll(n => n.IsContentControl || n.HasInlineContentControls());

    /// <summary>
    /// Gets all content controls of a specific type.
    /// </summary>
    public static IEnumerable<DocumentNode> GetContentControlsByType(this WordDocument document, ContentControlType type)
        => document.Root.GetContentControlsByType(type);

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
    public static DocumentNode? FindContentControlByTag(this WordDocument document, string tag)
        => document.Root.FindContentControlByTag(tag);

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
    public static DocumentNode? FindContentControlByAlias(this WordDocument document, string alias)
        => document.Root.FindContentControlByAlias(alias);

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
    public static DocumentNode? FindContentControlById(this WordDocument document, int id)
        => document.Root.FindContentControlById(id);

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
    public static bool SetContentControlValueByTag(this WordDocument document, string tag, string newValue)
        => document.Root.SetContentControlValueByTag(tag, newValue);

    /// <summary>
    /// Sets the value of a content control by tag.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="tag">The tag identifying the control</param>
    /// <param name="newValue">The value to set</param>
    /// <returns>True if control was found and updated</returns>
    public static bool SetContentControlValueByTag(this DocumentNode root, string tag, string newValue)
    {
        var control = root.FindContentControlByTag(tag);
        return control is not null && control.SetContentControlValue(newValue, props => props.Tag == tag);
    }

    /// <summary>
    /// Sets the value of a content control by alias.
    /// </summary>
    /// <returns>True if control was found and updated</returns>
    public static bool SetContentControlValueByAlias(this WordDocument document, string alias, string newValue)
        => document.Root.SetContentControlValueByAlias(alias, newValue);

    /// <summary>
    /// Sets the value of a content control by alias.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="alias">The alias identifying the control</param>
    /// <param name="newValue">The value to set</param>
    /// <returns>True if control was found and updated</returns>
    public static bool SetContentControlValueByAlias(this DocumentNode root, string alias, string newValue)
    {
        var control = root.FindContentControlByAlias(alias);
        return control is not null && control.SetContentControlValue(newValue, props => props.Alias == alias);
    }

    /// <summary>
    /// Sets a content control's value on the node that carries it.
    /// </summary>
    /// <param name="node">The node holding the control.</param>
    /// <param name="newValue">The value to set.</param>
    /// <param name="matches">Identifies which control on the node to update.</param>
    /// <returns>True when a matching control was updated.</returns>
    /// <remarks>
    /// <para>
    /// The edit is recorded so the writer applies it to the control inside the node's original XML;
    /// without the record the writer would emit the unmodified original and drop the new value.
    /// </para>
    /// <para>
    /// For an inline control the value replaces that control's runs only, leaving the text on either
    /// side of it alone.
    /// </para>
    /// </remarks>
    private static bool SetContentControlValue(
        this DocumentNode node, string newValue, Func<ContentControlProperties, bool> matches)
    {
        // Block-level control: the node itself carries the properties.
        if (node.ContentControlProperties is { } blockProps && matches(blockProps))
        {
            node.Text = newValue;
            blockProps.Value = newValue;

            if (node.Runs.Count > 0)
            {
                node.Runs.Clear();
                node.Runs.Add(new FormattedRun(newValue));
                node.MarkRunsChanged();
            }

            return true;
        }

        // Inline control: replace the runs belonging to that control, keeping the rest.
        var controlRuns = node.Runs
            .Where(r => r.ContentControlProperties is not null && matches(r.ContentControlProperties))
            .ToList();

        if (controlRuns.Count == 0) return false;

        controlRuns[0].Text = newValue;
        for (var i = 1; i < controlRuns.Count; i++)
        {
            controlRuns[i].Text = string.Empty;
        }

        if (controlRuns[0].ContentControlProperties is { } inlineProps)
        {
            inlineProps.Value = newValue;
        }

        node.MarkRunsChanged();
        return true;
    }

    /// <summary>
    /// Gets all content control tags in the document.
    /// </summary>
    public static IEnumerable<string> GetContentControlTags(this WordDocument document)
        => document.Root.GetContentControlTags();

    /// <summary>
    /// Gets all content control tags in the node tree.
    /// </summary>
    public static IEnumerable<string> GetContentControlTags(this DocumentNode root)
        => root.GetAllContentControls()
            .Where(n => !string.IsNullOrEmpty(n.ContentControlProperties?.Tag))
            .Select(n => n.ContentControlProperties!.Tag!);

    /// <summary>
    /// Gets content control properties by tag.
    /// </summary>
    public static ContentControlProperties? GetContentControlPropertiesByTag(this WordDocument document, string tag)
        => document.Root.GetContentControlPropertiesByTag(tag);

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

        // Block-level control: unwrap the SDT, keeping every block it wrapped.
        if (node.ContentControlProperties is not null &&
            (contentControlId is null || node.ContentControlProperties.Id == contentControlId) &&
            TryUnwrapSdt(node))
        {
            node.ContentControlProperties = null;
            node.Metadata.Remove(DocumentNode.IsSdtContentKey);
            node.Metadata.Remove(DocumentNode.IsSdtBlockKey);
            removed = true;
        }

        // Handle inline content controls in runs
        foreach (var run in node.Runs)
        {
            if (run.ContentControlProperties is not null &&
                (contentControlId is null || run.ContentControlProperties.Id == contentControlId))
            {
                run.ContentControlProperties = null;
                removed = true;
            }
        }

        if (removed)
        {
            node.MarkRunsChanged();
        }

        return removed;
    }

    /// <summary>
    /// Replaces a node's SDT XML with the content the SDT wrapped, keeping every block inside it.
    /// </summary>
    /// <param name="node">The node whose control is being removed.</param>
    /// <returns>False when the control could not be unwrapped without losing content.</returns>
    /// <remarks>
    /// The node's XML is kept rather than discarded, so the hyperlinks, fields, and formatting that
    /// sat inside the control survive its removal. A control wrapping several blocks already holds
    /// each of them as a child node, so it drops its own XML and lets those children be written in
    /// its place — they are no longer part of any control's content.
    /// </remarks>
    private static bool TryUnwrapSdt(DocumentNode node)
    {
        var originalXml = node.OriginalXml;

        if (string.IsNullOrEmpty(originalXml) ||
            !originalXml.TrimStart().StartsWith("<w:sdt", StringComparison.Ordinal))
        {
            ReleaseChildrenFromSdt(node);
            return true;
        }

        WP.SdtContentBlock? content;
        try
        {
            content = new WP.SdtBlock(originalXml).SdtContentBlock;
        }
        catch (Exception)
        {
            // Keep the control rather than losing its content to a parse failure.
            return false;
        }

        if (content is null) return false;

        var blocks = content.ChildElements
            .Where(child => child is not WP.SdtProperties and not WP.SdtEndCharProperties)
            .ToList();

        // Only children marked as physically inside the control count. A heading wrapped in a
        // control also gathers the rest of its section as children by the heading hierarchy, and
        // counting those made an ordinary heading control impossible to remove.
        var blocksHeldAsChildren = node.Children.Count(child => child.IsInsideParentSdt);

        // One paragraph, held by this node itself: keep the paragraph.
        if (blocks.Count == 1 && blocks[0] is WP.Paragraph paragraph && blocksHeldAsChildren == 0)
        {
            node.OriginalXml = paragraph.OuterXml;
            ReleaseChildrenFromSdt(node);
            return true;
        }

        // Several blocks: the node must already hold one child per block, or unwrapping would drop
        // whatever the tree does not represent.
        if (blocks.Count != blocksHeldAsChildren) return false;

        node.OriginalXml = null;
        ReleaseChildrenFromSdt(node);
        return true;
    }

    /// <summary>
    /// Clears the marker that tells the writer a child is already emitted with its parent's SDT XML.
    /// </summary>
    private static void ReleaseChildrenFromSdt(DocumentNode node)
    {
        foreach (var child in node.Children)
        {
            child.Metadata.Remove(DocumentNode.IsInsideParentSdtKey);
        }
    }

    /// <summary>
    /// Removes all content controls from the document, keeping text content.
    /// </summary>
    /// <returns>Number of nodes modified</returns>
    public static int RemoveAllContentControls(this WordDocument document)
        => document.Root.RemoveAllContentControls();

    /// <summary>
    /// Removes all content controls from the node tree, keeping text content.
    /// </summary>
    /// <returns>Number of nodes modified</returns>
    public static int RemoveAllContentControls(this DocumentNode root)
    {
        var count = 0;
        foreach (var node in root.FindAllContent(_ => true))
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
    public static bool RemoveContentControlByTag(this WordDocument document, string tag)
        => document.Root.RemoveContentControlByTag(tag);

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
    public static bool RemoveContentControlByAlias(this WordDocument document, string alias)
        => document.Root.RemoveContentControlByAlias(alias);

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
