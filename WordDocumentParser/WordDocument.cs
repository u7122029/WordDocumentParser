using System.Xml.Linq;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using WordDocumentParser.Models.Package;

namespace WordDocumentParser;

/// <summary>
/// Represents a complete Word document with its metadata, properties, styles, and content structure.
/// This class decouples document-level metadata from the content hierarchy.
/// </summary>
public class WordDocument
{
    private Dictionary<string, string>? _customProperties;
    private bool _customPropertiesParsed;

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

    #region Document Properties (Dictionary-Style Access)

    /// <summary>
    /// Gets or sets a document property by name.
    /// Supports core properties, extended properties, and custom properties.
    /// Property names are case-insensitive. Unknown property names are treated as custom properties.
    /// Set to null to delete a property.
    /// </summary>
    /// <param name="propertyName">The name of the property to get or set.</param>
    /// <returns>The property value as a string, or null if not found.</returns>
    public string? this[string propertyName]
    {
        get => GetProperty(propertyName);
        set
        {
            if (value is null)
                RemoveProperty(propertyName);
            else
                SetProperty(propertyName, value);
        }
    }

    /// <summary>
    /// Custom document properties dictionary.
    /// Changes to this dictionary are automatically serialized when the document is saved.
    /// </summary>
    public Dictionary<string, string> CustomProperties
    {
        get
        {
            EnsureCustomPropertiesParsed();
            return _customProperties!;
        }
    }

    /// <summary>
    /// Gets a document property value by name (case-insensitive).
    /// Checks core properties, extended properties, then custom properties.
    /// </summary>
    /// <param name="propertyName">The property name.</param>
    /// <returns>The property value as a string, or null if not found.</returns>
    public string? GetProperty(string propertyName)
    {
        var lowerName = propertyName.ToLowerInvariant();

        // Check core and extended properties first
        var builtInValue = lowerName switch
        {
            // Core properties
            "title" => CoreProperties?.Title,
            "subject" => CoreProperties?.Subject,
            "creator" or "author" => CoreProperties?.Creator,
            "keywords" => CoreProperties?.Keywords,
            "description" or "comments" => CoreProperties?.Description,
            "lastmodifiedby" => CoreProperties?.LastModifiedBy,
            "revision" => CoreProperties?.Revision,
            "created" => CoreProperties?.Created,
            "modified" => CoreProperties?.Modified,
            "category" => CoreProperties?.Category,
            "contentstatus" or "status" => CoreProperties?.ContentStatus,

            // Extended properties
            "template" => ExtendedProperties?.Template,
            "application" => ExtendedProperties?.Application,
            "appversion" => ExtendedProperties?.AppVersion,
            "company" => ExtendedProperties?.Company,
            "manager" => ExtendedProperties?.Manager,
            "pages" => ExtendedProperties?.Pages?.ToString(),
            "words" => ExtendedProperties?.Words?.ToString(),
            "characters" => ExtendedProperties?.Characters?.ToString(),
            "characterswithspaces" => ExtendedProperties?.CharactersWithSpaces?.ToString(),
            "lines" => ExtendedProperties?.Lines?.ToString(),
            "paragraphs" => ExtendedProperties?.Paragraphs?.ToString(),
            "totaltime" => ExtendedProperties?.TotalTime?.ToString(),

            _ => (string?)null
        };

        if (builtInValue is not null)
            return builtInValue;

        // Fall back to custom properties (case-insensitive lookup)
        EnsureCustomPropertiesParsed();
        var customKey = _customProperties!.Keys.FirstOrDefault(k =>
            string.Equals(k, propertyName, StringComparison.OrdinalIgnoreCase));
        return customKey is not null ? _customProperties[customKey] : null;
    }

    /// <summary>
    /// Sets a document property value by name (case-insensitive).
    /// Unknown property names are stored as custom properties.
    /// </summary>
    /// <param name="propertyName">The property name.</param>
    /// <param name="value">The value to set.</param>
    public void SetProperty(string propertyName, string? value)
    {
        if (value is null)
        {
            RemoveProperty(propertyName);
            return;
        }

        var lowerName = propertyName.ToLowerInvariant();

        if (DocumentPropertyHelpers.IsCoreProperty(lowerName))
        {
            PackageData.CoreProperties ??= new CoreProperties();
            SetCoreProperty(lowerName, value);
        }
        else if (DocumentPropertyHelpers.IsExtendedProperty(lowerName))
        {
            PackageData.ExtendedProperties ??= new ExtendedProperties();
            SetExtendedProperty(lowerName, value);
        }
        else
        {
            // Store as custom property (preserve original casing for new properties)
            EnsureCustomPropertiesParsed();
            var existingKey = _customProperties!.Keys.FirstOrDefault(k =>
                string.Equals(k, propertyName, StringComparison.OrdinalIgnoreCase));
            if (existingKey is not null)
                _customProperties[existingKey] = value;
            else
                _customProperties[propertyName] = value;
        }
    }

    /// <summary>
    /// Removes a document property by name (case-insensitive).
    /// </summary>
    /// <param name="propertyName">The property name to remove.</param>
    /// <returns>True if the property was found and removed.</returns>
    public bool RemoveProperty(string propertyName)
    {
        var lowerName = propertyName.ToLowerInvariant();

        if (DocumentPropertyHelpers.IsCoreProperty(lowerName))
        {
            if (CoreProperties is null) return false;
            return RemoveCoreProperty(lowerName);
        }

        if (DocumentPropertyHelpers.IsExtendedProperty(lowerName))
        {
            if (ExtendedProperties is null) return false;
            return RemoveExtendedProperty(lowerName);
        }

        // Remove from custom properties
        EnsureCustomPropertiesParsed();
        var existingKey = _customProperties!.Keys.FirstOrDefault(k =>
            string.Equals(k, propertyName, StringComparison.OrdinalIgnoreCase));
        return existingKey is not null && _customProperties.Remove(existingKey);
    }

    /// <summary>
    /// Checks if a property with the given name exists and has a value.
    /// </summary>
    /// <param name="propertyName">The property name (case-insensitive).</param>
    /// <returns>True if the property exists and has a non-null value.</returns>
    public bool HasProperty(string propertyName) => GetProperty(propertyName) is not null;

    /// <summary>
    /// Gets all properties (core, extended, and custom) that have values.
    /// </summary>
    /// <returns>Dictionary of property names and their values.</returns>
    public Dictionary<string, string> GetAllProperties()
    {
        var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Add built-in properties
        foreach (var name in BuiltInPropertyNames)
        {
            var value = GetProperty(name);
            if (value is not null)
                properties[name] = value;
        }

        // Add custom properties
        foreach (var kvp in CustomProperties)
        {
            properties[kvp.Key] = kvp.Value;
        }

        return properties;
    }

    /// <summary>
    /// Gets all built-in property names (core and extended).
    /// </summary>
    public static IReadOnlyList<string> BuiltInPropertyNames { get; } =
    [
        // Core properties
        "Title", "Subject", "Creator", "Keywords", "Description",
        "LastModifiedBy", "Revision", "Created", "Modified", "Category", "ContentStatus",
        // Extended properties
        "Template", "Application", "AppVersion", "Company", "Manager",
        "Pages", "Words", "Characters", "CharactersWithSpaces", "Lines", "Paragraphs", "TotalTime"
    ];

    #endregion

    #region Custom Properties XML Serialization

    private void EnsureCustomPropertiesParsed()
    {
        if (_customPropertiesParsed) return;

        _customProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        _customPropertiesParsed = true;

        if (string.IsNullOrEmpty(PackageData.CustomPropertiesXml)) return;

        try
        {
            var doc = XDocument.Parse(PackageData.CustomPropertiesXml);

            // Find all property elements regardless of namespace
            // The element local name is "property" in Open XML custom properties
            foreach (var prop in doc.Descendants().Where(e => e.Name.LocalName == "property"))
            {
                var name = prop.Attribute("name")?.Value;
                if (name is not null)
                {
                    // Get the value from any child element (lpwstr, i4, bool, filetime, etc.)
                    var valueElement = prop.Elements().FirstOrDefault();
                    var value = valueElement?.Value ?? string.Empty;
                    _customProperties[name] = value;
                }
            }
        }
        catch
        {
            // If parsing fails, start with empty dictionary
        }
    }

    /// <summary>
    /// Serializes the custom properties back to XML format for saving.
    /// Called automatically by the writer.
    /// </summary>
    internal void SyncCustomPropertiesToXml()
    {
        if (!_customPropertiesParsed || _customProperties is null || _customProperties.Count == 0)
        {
            if (_customPropertiesParsed && (_customProperties is null || _customProperties.Count == 0))
                PackageData.CustomPropertiesXml = null;
            return;
        }

        XNamespace ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        var props = new XElement(ns + "Properties",
            new XAttribute(XNamespace.Xmlns + "vt", vt));

        var pid = 2; // Property IDs start at 2
        foreach (var kvp in _customProperties)
        {
            var prop = new XElement(ns + "property",
                new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                new XAttribute("pid", pid++),
                new XAttribute("name", kvp.Key),
                new XElement(vt + "lpwstr", kvp.Value));
            props.Add(prop);
        }

        PackageData.CustomPropertiesXml = props.ToString();
    }

    #endregion

    #region Private Property Setters/Removers

    private void SetCoreProperty(string lowerName, string value)
    {
        switch (lowerName)
        {
            case "title": CoreProperties!.Title = value; break;
            case "subject": CoreProperties!.Subject = value; break;
            case "creator": case "author": CoreProperties!.Creator = value; break;
            case "keywords": CoreProperties!.Keywords = value; break;
            case "description": case "comments": CoreProperties!.Description = value; break;
            case "lastmodifiedby": CoreProperties!.LastModifiedBy = value; break;
            case "revision": CoreProperties!.Revision = value; break;
            case "created": CoreProperties!.Created = value; break;
            case "modified": CoreProperties!.Modified = value; break;
            case "category": CoreProperties!.Category = value; break;
            case "contentstatus": case "status": CoreProperties!.ContentStatus = value; break;
        }
    }

    private void SetExtendedProperty(string lowerName, string value)
    {
        switch (lowerName)
        {
            case "template": ExtendedProperties!.Template = value; break;
            case "application": ExtendedProperties!.Application = value; break;
            case "appversion": ExtendedProperties!.AppVersion = value; break;
            case "company": ExtendedProperties!.Company = value; break;
            case "manager": ExtendedProperties!.Manager = value; break;
            case "pages": ExtendedProperties!.Pages = int.TryParse(value, out var p) ? p : null; break;
            case "words": ExtendedProperties!.Words = int.TryParse(value, out var w) ? w : null; break;
            case "characters": ExtendedProperties!.Characters = int.TryParse(value, out var c) ? c : null; break;
            case "characterswithspaces": ExtendedProperties!.CharactersWithSpaces = int.TryParse(value, out var cs) ? cs : null; break;
            case "lines": ExtendedProperties!.Lines = int.TryParse(value, out var l) ? l : null; break;
            case "paragraphs": ExtendedProperties!.Paragraphs = int.TryParse(value, out var pg) ? pg : null; break;
            case "totaltime": ExtendedProperties!.TotalTime = int.TryParse(value, out var t) ? t : null; break;
        }
    }

    private bool RemoveCoreProperty(string lowerName)
    {
        var hadValue = GetProperty(lowerName) is not null;
        switch (lowerName)
        {
            case "title": CoreProperties!.Title = null; break;
            case "subject": CoreProperties!.Subject = null; break;
            case "creator": case "author": CoreProperties!.Creator = null; break;
            case "keywords": CoreProperties!.Keywords = null; break;
            case "description": case "comments": CoreProperties!.Description = null; break;
            case "lastmodifiedby": CoreProperties!.LastModifiedBy = null; break;
            case "revision": CoreProperties!.Revision = null; break;
            case "created": CoreProperties!.Created = null; break;
            case "modified": CoreProperties!.Modified = null; break;
            case "category": CoreProperties!.Category = null; break;
            case "contentstatus": case "status": CoreProperties!.ContentStatus = null; break;
        }
        return hadValue;
    }

    private bool RemoveExtendedProperty(string lowerName)
    {
        var hadValue = GetProperty(lowerName) is not null;
        switch (lowerName)
        {
            case "template": ExtendedProperties!.Template = null; break;
            case "application": ExtendedProperties!.Application = null; break;
            case "appversion": ExtendedProperties!.AppVersion = null; break;
            case "company": ExtendedProperties!.Company = null; break;
            case "manager": ExtendedProperties!.Manager = null; break;
            case "pages": ExtendedProperties!.Pages = null; break;
            case "words": ExtendedProperties!.Words = null; break;
            case "characters": ExtendedProperties!.Characters = null; break;
            case "characterswithspaces": ExtendedProperties!.CharactersWithSpaces = null; break;
            case "lines": ExtendedProperties!.Lines = null; break;
            case "paragraphs": ExtendedProperties!.Paragraphs = null; break;
            case "totaltime": ExtendedProperties!.TotalTime = null; break;
        }
        return hadValue;
    }

    #endregion

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
    /// </summary>
    public string? StylesXml
    {
        get => PackageData.StylesXml;
        set => PackageData.StylesXml = value;
    }

    /// <summary>
    /// The theme XML content (theme/theme1.xml).
    /// </summary>
    public string? ThemeXml
    {
        get => PackageData.ThemeXml;
        set => PackageData.ThemeXml = value;
    }

    /// <summary>
    /// The font table XML content (fontTable.xml).
    /// </summary>
    public string? FontTableXml
    {
        get => PackageData.FontTableXml;
        set => PackageData.FontTableXml = value;
    }

    /// <summary>
    /// The numbering definitions XML content (numbering.xml).
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

    #endregion

    /// <summary>
    /// Returns a tree representation of the document structure.
    /// </summary>
    public string ToTreeString() => Root.ToTreeString();

    /// <summary>
    /// Returns a string representation of this document.
    /// </summary>
    public override string ToString()
    {
        var title = Title ?? FileName ?? "Untitled";
        var nodeCount = Root.FindAll(_ => true).Count();
        return $"WordDocument: {title} ({nodeCount} nodes)";
    }
}
