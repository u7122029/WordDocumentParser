using System.Collections.Generic;

namespace WordDocumentParser
{
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
        public Dictionary<string, string> Headers { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Footer parts - key is the relationship ID, value is the XML content
        /// </summary>
        public Dictionary<string, string> Footers { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Image parts - key is the relationship ID, value is the image data
        /// </summary>
        public Dictionary<string, ImagePartData> Images { get; set; } = new Dictionary<string, ImagePartData>();

        /// <summary>
        /// Section properties from the original document
        /// </summary>
        public List<string> SectionPropertiesXml { get; set; } = new List<string>();

        /// <summary>
        /// The original document.xml content (for reference)
        /// </summary>
        public string? OriginalDocumentXml { get; set; }

        /// <summary>
        /// Custom XML parts - key is the part URI, value is the XML content
        /// </summary>
        public Dictionary<string, CustomXmlPartData> CustomXmlParts { get; set; } = new Dictionary<string, CustomXmlPartData>();

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
        public Dictionary<string, HyperlinkRelationshipData> HyperlinkRelationships { get; set; } = new Dictionary<string, HyperlinkRelationshipData>();

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
        public Dictionary<string, ImagePartData> GlossaryImages { get; set; } = new Dictionary<string, ImagePartData>();
    }

    /// <summary>
    /// Stores hyperlink relationship data
    /// </summary>
    public class HyperlinkRelationshipData
    {
        public string Url { get; set; } = string.Empty;
        public bool IsExternal { get; set; } = true;
    }

    /// <summary>
    /// Stores custom XML part data
    /// </summary>
    public class CustomXmlPartData
    {
        public string XmlContent { get; set; } = string.Empty;
        public string? PropertiesXml { get; set; }
    }

    /// <summary>
    /// Core document properties
    /// </summary>
    public class CoreProperties
    {
        public string? Title { get; set; }
        public string? Subject { get; set; }
        public string? Creator { get; set; }
        public string? Keywords { get; set; }
        public string? Description { get; set; }
        public string? LastModifiedBy { get; set; }
        public string? Revision { get; set; }
        public string? Created { get; set; }
        public string? Modified { get; set; }
        public string? Category { get; set; }
        public string? ContentStatus { get; set; }
    }

    /// <summary>
    /// Extended document properties
    /// </summary>
    public class ExtendedProperties
    {
        public string? Template { get; set; }
        public string? Application { get; set; }
        public string? AppVersion { get; set; }
        public string? Company { get; set; }
        public int? Pages { get; set; }
        public int? Words { get; set; }
        public int? Characters { get; set; }
        public int? CharactersWithSpaces { get; set; }
        public int? Lines { get; set; }
        public int? Paragraphs { get; set; }
        public string? Manager { get; set; }
        public int? TotalTime { get; set; }
    }

    /// <summary>
    /// Stores image part data
    /// </summary>
    public class ImagePartData
    {
        public string ContentType { get; set; } = string.Empty;
        public byte[] Data { get; set; } = Array.Empty<byte>();
        public string OriginalRelationshipId { get; set; } = string.Empty;
        public string? OriginalUri { get; set; }
    }
}
