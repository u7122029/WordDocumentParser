using WordDocumentParser.Models.ContentControls;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for working with document property fields.
/// </summary>
public static class DocumentPropertyExtensions
{
    /// <summary>
    /// Gets all document property content controls.
    /// </summary>
    public static IEnumerable<DocumentNode> GetDocumentPropertyControls(this WordDocument document)
        => document.Root.GetDocumentPropertyControls();

    /// <summary>
    /// Gets all document property content controls.
    /// </summary>
    public static IEnumerable<DocumentNode> GetDocumentPropertyControls(this DocumentNode root)
        => root.GetContentControlsByType(ContentControlType.DocumentProperty);

    /// <summary>
    /// Gets all nodes that contain document property fields.
    /// </summary>
    public static IEnumerable<DocumentNode> GetNodesWithDocumentPropertyFields(this DocumentNode root)
        => root.FindAll(n => n.HasDocumentPropertyFields);

    /// <summary>
    /// Gets all document property fields in the document.
    /// </summary>
    public static IEnumerable<DocumentPropertyField> GetAllDocumentPropertyFields(this WordDocument document)
        => document.Root.GetAllDocumentPropertyFields();

    /// <summary>
    /// Gets all document property fields in the document.
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
    /// Gets all text with metadata annotations for document properties and content controls.
    /// </summary>
    public static string GetAllTextWithMetadata(this WordDocument document)
        => document.Root.GetAllTextWithMetadata();

    /// <summary>
    /// Gets all text with metadata annotations for document properties and content controls.
    /// </summary>
    public static string GetAllTextWithMetadata(this DocumentNode node)
    {
        var texts = new List<string>();

        if (node.Type is not Core.ContentType.Table and not Core.ContentType.Image)
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
    /// Removes a document property field from a node's runs, keeping text content.
    /// </summary>
    /// <param name="node">The node containing the field</param>
    /// <param name="propertyName">Property name to remove, or null for all</param>
    /// <returns>True if any field was removed</returns>
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

        if (removed)
        {
            node.OriginalXml = null;
        }

        return removed;
    }

    /// <summary>
    /// Removes all document property fields from a document, keeping text content.
    /// </summary>
    /// <returns>Number of nodes modified</returns>
    public static int RemoveAllDocumentPropertyFields(this WordDocument document)
        => document.Root.RemoveAllDocumentPropertyFields();

    /// <summary>
    /// Removes all document property fields from a document, keeping text content.
    /// </summary>
    /// <returns>Number of nodes modified</returns>
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
}
