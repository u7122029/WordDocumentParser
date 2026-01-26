using WordDocumentParser.Core;
using WordDocumentParser.Models.Package;

namespace WordDocumentParser;

/// <summary>
/// Represents a complete Word document with its metadata, properties, styles, and content structure.
/// This class decouples document-level metadata from the content hierarchy.
/// </summary>
public class WordDocument
{
    /// <summary>
    /// The root node of the document content structure.
    /// Contains the hierarchical tree of headings, paragraphs, tables, and other content.
    /// </summary>
    public DocumentNode Root { get; set; }

    /// <summary>
    /// The original file name of the document (if available).
    /// </summary>
    public string? FileName { get; set; }

    /// <summary>
    /// Complete package data for round-trip fidelity.
    /// Contains all XML parts, styles, themes, and resources from the original document.
    /// </summary>
    public DocumentPackageData PackageData { get; set; } = new();

    /// <summary>
    /// Creates a new empty document with a root Document node.
    /// </summary>
    public WordDocument()
    {
        Root = new DocumentNode(ContentType.Document);
    }

    /// <summary>
    /// Creates a document with the specified root node.
    /// </summary>
    /// <param name="root">The root node of the document structure.</param>
    public WordDocument(DocumentNode root)
    {
        Root = root;
    }

    /// <summary>
    /// Creates a document with the specified root node and package data.
    /// </summary>
    /// <param name="root">The root node of the document structure.</param>
    /// <param name="packageData">The document package data for round-trip fidelity.</param>
    public WordDocument(DocumentNode root, DocumentPackageData packageData)
    {
        Root = root;
        PackageData = packageData;
    }

    #region Document Properties (Convenience Accessors)

    /// <summary>
    /// Core document properties (title, author, subject, etc.).
    /// </summary>
    public CoreProperties? CoreProperties
    {
        get => PackageData.CoreProperties;
        set => PackageData.CoreProperties = value;
    }

    /// <summary>
    /// Extended document properties (company, template, word count, etc.).
    /// </summary>
    public ExtendedProperties? ExtendedProperties
    {
        get => PackageData.ExtendedProperties;
        set => PackageData.ExtendedProperties = value;
    }

    /// <summary>
    /// Gets the document title from core properties.
    /// </summary>
    public string? Title => CoreProperties?.Title;

    /// <summary>
    /// Gets the document author/creator from core properties.
    /// </summary>
    public string? Author => CoreProperties?.Creator;

    /// <summary>
    /// Gets the document subject from core properties.
    /// </summary>
    public string? Subject => CoreProperties?.Subject;

    /// <summary>
    /// Gets the document description from core properties.
    /// </summary>
    public string? Description => CoreProperties?.Description;

    /// <summary>
    /// Gets the document keywords from core properties.
    /// </summary>
    public string? Keywords => CoreProperties?.Keywords;

    /// <summary>
    /// Gets the document category from core properties.
    /// </summary>
    public string? Category => CoreProperties?.Category;

    /// <summary>
    /// Gets the document creation date from core properties (as ISO 8601 string).
    /// </summary>
    public string? Created => CoreProperties?.Created;

    /// <summary>
    /// Gets the document last modified date from core properties (as ISO 8601 string).
    /// </summary>
    public string? Modified => CoreProperties?.Modified;

    /// <summary>
    /// Gets the company name from extended properties.
    /// </summary>
    public string? Company => ExtendedProperties?.Company;

    /// <summary>
    /// Gets the application that created/modified the document from extended properties.
    /// </summary>
    public string? Application => ExtendedProperties?.Application;

    /// <summary>
    /// Gets the document template from extended properties.
    /// </summary>
    public string? Template => ExtendedProperties?.Template;

    #endregion

    #region Style and Formatting Resources

    /// <summary>
    /// The styles XML content (styles.xml).
    /// Contains all style definitions used in the document.
    /// </summary>
    public string? StylesXml
    {
        get => PackageData.StylesXml;
        set => PackageData.StylesXml = value;
    }

    /// <summary>
    /// The theme XML content (theme/theme1.xml).
    /// Contains theme colors, fonts, and effects.
    /// </summary>
    public string? ThemeXml
    {
        get => PackageData.ThemeXml;
        set => PackageData.ThemeXml = value;
    }

    /// <summary>
    /// The font table XML content (fontTable.xml).
    /// Contains font definitions and substitutions.
    /// </summary>
    public string? FontTableXml
    {
        get => PackageData.FontTableXml;
        set => PackageData.FontTableXml = value;
    }

    /// <summary>
    /// The numbering definitions XML content (numbering.xml).
    /// Contains list formatting definitions.
    /// </summary>
    public string? NumberingXml
    {
        get => PackageData.NumberingXml;
        set => PackageData.NumberingXml = value;
    }

    #endregion

    #region Document Parts

    /// <summary>
    /// Header parts - key is the relationship ID, value is the XML content.
    /// </summary>
    public Dictionary<string, string> Headers => PackageData.Headers;

    /// <summary>
    /// Footer parts - key is the relationship ID, value is the XML content.
    /// </summary>
    public Dictionary<string, string> Footers => PackageData.Footers;

    /// <summary>
    /// Image parts - key is the relationship ID, value is the image data.
    /// </summary>
    public Dictionary<string, ImagePartData> Images => PackageData.Images;

    /// <summary>
    /// Hyperlink relationships - key is the relationship ID, value is the hyperlink data.
    /// </summary>
    public Dictionary<string, HyperlinkRelationshipData> HyperlinkRelationships => PackageData.HyperlinkRelationships;

    /// <summary>
    /// Custom XML parts - key is the part URI, value is the XML content.
    /// </summary>
    public Dictionary<string, CustomXmlPartData> CustomXmlParts => PackageData.CustomXmlParts;

    #endregion

    #region Content Traversal

    /// <summary>
    /// Gets all nodes in the document tree (including the root).
    /// </summary>
    public IEnumerable<DocumentNode> GetAllNodes()
    {
        yield return Root;
        foreach (var descendant in Root.GetDescendants())
        {
            yield return descendant;
        }
    }

    /// <summary>
    /// Gets all heading nodes in document order.
    /// </summary>
    public IEnumerable<DocumentNode> GetHeadings()
    {
        return GetAllNodes().Where(n => n.Type == ContentType.Heading);
    }

    /// <summary>
    /// Gets all paragraph nodes in document order.
    /// </summary>
    public IEnumerable<DocumentNode> GetParagraphs()
    {
        return GetAllNodes().Where(n => n.Type == ContentType.Paragraph);
    }

    /// <summary>
    /// Gets all table nodes in document order.
    /// </summary>
    public IEnumerable<DocumentNode> GetTables()
    {
        return GetAllNodes().Where(n => n.Type == ContentType.Table);
    }

    /// <summary>
    /// Gets all content control nodes in document order.
    /// </summary>
    public IEnumerable<DocumentNode> GetContentControls()
    {
        return GetAllNodes().Where(n => n.IsContentControl);
    }

    /// <summary>
    /// Gets all nodes that contain document property fields.
    /// </summary>
    public IEnumerable<DocumentNode> GetNodesWithDocumentPropertyFields()
    {
        return GetAllNodes().Where(n => n.HasDocumentPropertyFields);
    }

    #endregion

    /// <summary>
    /// Returns a tree representation of the document structure.
    /// </summary>
    public string ToTreeString()
    {
        return Root.ToTreeString();
    }

    /// <summary>
    /// Returns a string representation of this document.
    /// </summary>
    public override string ToString()
    {
        var title = Title ?? FileName ?? "Untitled";
        var nodeCount = GetAllNodes().Count();
        return $"WordDocument: {title} ({nodeCount} nodes)";
    }
}
