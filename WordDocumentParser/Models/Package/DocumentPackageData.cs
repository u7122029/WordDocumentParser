namespace WordDocumentParser.Models.Package;

/// <summary>
/// Stores the original document package data for round-trip fidelity.
/// This preserves styles, themes, fonts, properties, and other document parts.
/// </summary>
public class DocumentPackageData
{
    /// <summary>
    /// Original styles.xml content
    /// </summary>
    public string? StylesXml { get; set; }

    /// <summary>
    /// Original theme XML content (theme/theme1.xml)
    /// </summary>
    public string? ThemeXml { get; set; }

    /// <summary>
    /// Original font table XML (fontTable.xml)
    /// </summary>
    public string? FontTableXml { get; set; }

    /// <summary>
    /// Original numbering definitions XML (numbering.xml)
    /// </summary>
    public string? NumberingXml { get; set; }

    /// <summary>
    /// Original document settings XML (settings.xml)
    /// </summary>
    public string? SettingsXml { get; set; }

    /// <summary>
    /// Original web settings XML (webSettings.xml)
    /// </summary>
    public string? WebSettingsXml { get; set; }

    /// <summary>
    /// Original footnotes XML
    /// </summary>
    public string? FootnotesXml { get; set; }

    /// <summary>
    /// Original endnotes XML
    /// </summary>
    public string? EndnotesXml { get; set; }

    /// <summary>
    /// Core document properties (author, title, etc.)
    /// </summary>
    public CoreProperties? CoreProperties { get; set; }

    /// <summary>
    /// Extended document properties (company, template, word count, etc.)
    /// </summary>
    public ExtendedProperties? ExtendedProperties { get; set; }

    /// <summary>
    /// Custom document properties
    /// </summary>
    public string? CustomPropertiesXml { get; set; }

    /// <summary>
    /// Header parts - key is the relationship ID, value is the XML content
    /// </summary>
    public Dictionary<string, string> Headers { get; set; } = [];

    /// <summary>
    /// Footer parts - key is the relationship ID, value is the XML content
    /// </summary>
    public Dictionary<string, string> Footers { get; set; } = [];

    /// <summary>
    /// Image parts - key is the relationship ID, value is the image data
    /// </summary>
    public Dictionary<string, ImagePartData> Images { get; set; } = [];

    /// <summary>
    /// Section properties from the original document
    /// </summary>
    public List<string> SectionPropertiesXml { get; set; } = [];

    /// <summary>
    /// The original document.xml content (for reference)
    /// </summary>
    public string? OriginalDocumentXml { get; set; }

    /// <summary>
    /// Custom XML parts - key is the part URI, value is the XML content
    /// </summary>
    public Dictionary<string, CustomXmlPartData> CustomXmlParts { get; set; } = [];

    /// <summary>
    /// Original core.xml content for exact round-trip
    /// </summary>
    public string? CorePropertiesXml { get; set; }

    /// <summary>
    /// Original app.xml content for exact round-trip
    /// </summary>
    public string? AppPropertiesXml { get; set; }

    /// <summary>
    /// Hyperlink relationships - key is the relationship ID, value is the URL
    /// </summary>
    public Dictionary<string, HyperlinkRelationshipData> HyperlinkRelationships { get; set; } = [];

    /// <summary>
    /// Glossary document XML (for Quick Parts, building blocks, document property fields)
    /// </summary>
    public string? GlossaryDocumentXml { get; set; }

    /// <summary>
    /// Glossary document styles XML
    /// </summary>
    public string? GlossaryStylesXml { get; set; }

    /// <summary>
    /// Glossary document fonts XML
    /// </summary>
    public string? GlossaryFontTableXml { get; set; }

    /// <summary>
    /// Images from glossary document part - key is relationship ID, value is image data
    /// </summary>
    public Dictionary<string, ImagePartData> GlossaryImages { get; set; } = [];
}