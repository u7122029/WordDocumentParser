using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Models.ContentControls;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Images;
using WordDocumentParser.Models.Package;
using WordDocumentParser.Models.Tables;
using WordDocumentParser.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordDocumentParser;

/// <summary>
/// Writes a document tree structure to a Word document (.docx file).
/// Preserves all formatting for round-trip fidelity by restoring original document parts.
/// </summary>
public class WordDocumentTreeWriter : IDocumentWriter
{
    private WordprocessingDocument? _document;
    private MainDocumentPart? _mainPart;
    private Body? _body;
    private NumberingDefinitionsPart? _numberingPart;
    private int _currentListId = 1;
    private uint _imageCounter = 1;
    private readonly Dictionary<string, string> _hyperlinkRelationships = [];
    private readonly Dictionary<string, string> _imageRelationshipMapping = [];
    private readonly Dictionary<string, string> _hyperlinkRelationshipMapping = [];
    private readonly Dictionary<string, string> _headerRelationshipMapping = [];
    private readonly Dictionary<string, string> _footerRelationshipMapping = [];
    private DocumentPackageData? _packageData;

    /// <summary>
    /// Writes a document to a file
    /// </summary>
    public void WriteToFile(WordDocument document, string filePath)
    {
        _document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
        try
        {
            BuildDocument(document);
            _document.Save();
        }
        finally
        {
            _document.Dispose();
            _document = null;
        }
    }

    /// <summary>
    /// Writes a document to a stream
    /// </summary>
    public void WriteToStream(WordDocument document, Stream stream)
    {
        _document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, false);
        try
        {
            BuildDocument(document);
            _document.Save();
        }
        finally
        {
            _document.Dispose();
            _document = null;
        }
    }

    /// <summary>
    /// Builds the document content from the WordDocument
    /// </summary>
    private void BuildDocument(WordDocument document)
    {
        // Sync custom properties to XML before saving
        document.SyncCustomPropertiesToXml();

        // Clear mappings for fresh document
        _imageRelationshipMapping.Clear();
        _hyperlinkRelationshipMapping.Clear();
        _hyperlinkRelationships.Clear();
        _headerRelationshipMapping.Clear();
        _footerRelationshipMapping.Clear();

        _packageData = document.PackageData;
        _mainPart = _document!.AddMainDocumentPart();
        _mainPart.Document = new Document();
        _body = new Body();
        _mainPart.Document.Append(_body);

        // Restore original document parts if available, otherwise use defaults
        if (_packageData != null)
        {
            RestoreDocumentParts();
        }
        else
        {
            AddStyleDefinitions();
        }

        // Process the document tree
        ProcessNode(document.Root);

        // Add section properties for page layout
        AddSectionProperties();

        // Restore document properties
        if (_packageData != null)
        {
            RestoreDocumentProperties();
        }
    }

    /// <summary>
    /// Restores all original document parts from package data
    /// </summary>
    private void RestoreDocumentParts()
    {
        // Restore styles (clean XML attributes)
        if (!string.IsNullOrEmpty(_packageData!.StylesXml))
        {
            var stylesPart = _mainPart!.AddNewPart<StyleDefinitionsPart>();
            var cleanedStylesXml = CleanXmlAttributes(_packageData.StylesXml);
            stylesPart.Styles = new Styles(cleanedStylesXml);
            // Fix any remaining indentation attributes in styles
            FixIndentationAttributes(stylesPart.Styles);
        }
        else
        {
            AddStyleDefinitions();
        }

        // Restore theme
        if (!string.IsNullOrEmpty(_packageData.ThemeXml))
        {
            var themePart = _mainPart!.AddNewPart<ThemePart>();
            var cleanedThemeXml = CleanXmlAttributes(_packageData.ThemeXml);
            themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(cleanedThemeXml);
        }

        // Restore font table
        if (!string.IsNullOrEmpty(_packageData.FontTableXml))
        {
            var fontTablePart = _mainPart!.AddNewPart<FontTablePart>();
            var cleanedFontTableXml = CleanXmlAttributes(_packageData.FontTableXml);
            fontTablePart.Fonts = new Fonts(cleanedFontTableXml);
        }

        // Restore numbering definitions (clean XML attributes - this is where w:start errors come from)
        if (!string.IsNullOrEmpty(_packageData.NumberingXml))
        {
            _numberingPart = _mainPart!.AddNewPart<NumberingDefinitionsPart>();
            var cleanedNumberingXml = CleanXmlAttributes(_packageData.NumberingXml);
            _numberingPart.Numbering = new Numbering(cleanedNumberingXml);
            // Fix any remaining indentation attributes
            FixIndentationAttributes(_numberingPart.Numbering);
        }

        // Restore document settings (clean XML attributes)
        if (!string.IsNullOrEmpty(_packageData.SettingsXml))
        {
            var settingsPart = _mainPart!.AddNewPart<DocumentSettingsPart>();
            var cleanedSettingsXml = CleanXmlAttributes(_packageData.SettingsXml);
            settingsPart.Settings = new Settings(cleanedSettingsXml);
        }

        // Restore web settings
        if (!string.IsNullOrEmpty(_packageData.WebSettingsXml))
        {
            var webSettingsPart = _mainPart!.AddNewPart<WebSettingsPart>();
            var cleanedWebSettingsXml = CleanXmlAttributes(_packageData.WebSettingsXml);
            webSettingsPart.WebSettings = new WebSettings(cleanedWebSettingsXml);
        }

        // Restore footnotes
        if (!string.IsNullOrEmpty(_packageData.FootnotesXml))
        {
            var footnotesPart = _mainPart!.AddNewPart<FootnotesPart>();
            var cleanedFootnotesXml = CleanXmlAttributes(_packageData.FootnotesXml);
            footnotesPart.Footnotes = new Footnotes(cleanedFootnotesXml);
        }

        // Restore endnotes
        if (!string.IsNullOrEmpty(_packageData.EndnotesXml))
        {
            var endnotesPart = _mainPart!.AddNewPart<EndnotesPart>();
            var cleanedEndnotesXml = CleanXmlAttributes(_packageData.EndnotesXml);
            endnotesPart.Endnotes = new Endnotes(cleanedEndnotesXml);
        }

        // Restore images and create relationship mapping
        foreach (var kvp in _packageData.Images)
        {
            var originalRelId = kvp.Key;
            var imageData = kvp.Value;

            var imagePart = _mainPart!.AddImagePart(imageData.ContentType);
            using (var stream = new MemoryStream(imageData.Data))
            {
                imagePart.FeedData(stream);
            }

            var newRelId = _mainPart.GetIdOfPart(imagePart);
            _imageRelationshipMapping[originalRelId] = newRelId;
        }

        // Restore hyperlink relationships and create mapping
        foreach (var kvp in _packageData.HyperlinkRelationships)
        {
            var originalRelId = kvp.Key;
            var hyperlinkData = kvp.Value;

            try
            {
                var uri = new Uri(hyperlinkData.Url, UriKind.RelativeOrAbsolute);
                var rel = _mainPart!.AddHyperlinkRelationship(uri, hyperlinkData.IsExternal);
                _hyperlinkRelationshipMapping[originalRelId] = rel.Id;
            }
            catch
            {
                // Skip invalid hyperlinks
            }
        }

        // Restore headers and create relationship mapping
        // Note: Headers/footers with complex content (SDT blocks, field codes) may have
        // some content simplified by the OpenXML SDK during parsing. This is a known
        // limitation of the SDK's typed API.
        foreach (var kvp in _packageData.Headers)
        {
            var originalRelId = kvp.Key;
            var headerPart = _mainPart!.AddNewPart<HeaderPart>();

            // Use the original XML as-is (don't clean) to preserve all content
            var headerXml = kvp.Value;

            // Collect image relationship mappings (we'll add images after loading XML)
            var headerImageMapping = new Dictionary<string, string>();

            // If there are images in this header, we need to add them and update IDs
            if (_packageData.HeaderImages.TryGetValue(originalRelId, out var headerImages))
            {
                foreach (var imgKvp in headerImages)
                {
                    var origImgRelId = imgKvp.Key;
                    var imageData = imgKvp.Value;

                    var imagePart = headerPart.AddImagePart(imageData.ContentType);
                    using (var stream = new MemoryStream(imageData.Data))
                    {
                        imagePart.FeedData(stream);
                    }
                    var newImgRelId = headerPart.GetIdOfPart(imagePart);
                    headerImageMapping[origImgRelId] = newImgRelId;
                }

                // Update XML with new image relationship IDs
                foreach (var imgMap in headerImageMapping)
                {
                    headerXml = headerXml.Replace($"r:embed=\"{imgMap.Key}\"", $"r:embed=\"{imgMap.Value}\"");
                    headerXml = headerXml.Replace($"r:id=\"{imgMap.Key}\"", $"r:id=\"{imgMap.Value}\"");
                }
            }

            // Create Header from XML - the SDK will parse and preserve recognized content
            headerPart.Header = new Header(headerXml);

            var newRelId = _mainPart.GetIdOfPart(headerPart);
            _headerRelationshipMapping[originalRelId] = newRelId;
        }

        // Restore footers and create relationship mapping
        foreach (var kvp in _packageData.Footers)
        {
            var originalRelId = kvp.Key;
            var footerPart = _mainPart!.AddNewPart<FooterPart>();

            // Use the original XML as-is (don't clean) to preserve all content
            var footerXml = kvp.Value;

            // Collect image relationship mappings
            var footerImageMapping = new Dictionary<string, string>();

            // If there are images in this footer, add them and update IDs
            if (_packageData.FooterImages.TryGetValue(originalRelId, out var footerImages))
            {
                foreach (var imgKvp in footerImages)
                {
                    var origImgRelId = imgKvp.Key;
                    var imageData = imgKvp.Value;

                    var imagePart = footerPart.AddImagePart(imageData.ContentType);
                    using (var stream = new MemoryStream(imageData.Data))
                    {
                        imagePart.FeedData(stream);
                    }
                    var newImgRelId = footerPart.GetIdOfPart(imagePart);
                    footerImageMapping[origImgRelId] = newImgRelId;
                }

                // Update XML with new image relationship IDs
                foreach (var imgMap in footerImageMapping)
                {
                    footerXml = footerXml.Replace($"r:embed=\"{imgMap.Key}\"", $"r:embed=\"{imgMap.Value}\"");
                    footerXml = footerXml.Replace($"r:id=\"{imgMap.Key}\"", $"r:id=\"{imgMap.Value}\"");
                }
            }

            // Create Footer from XML - the SDK will parse and preserve recognized content
            footerPart.Footer = new Footer(footerXml);

            var newRelId = _mainPart.GetIdOfPart(footerPart);
            _footerRelationshipMapping[originalRelId] = newRelId;
        }

        // Restore custom XML parts
        foreach (var kvp in _packageData.CustomXmlParts)
        {
            try
            {
                var customXmlPart = _mainPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using (var stream = customXmlPart.GetStream(FileMode.Create))
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write(kvp.Value.XmlContent);
                }

                // Add properties part if available
                if (!string.IsNullOrEmpty(kvp.Value.PropertiesXml))
                {
                    var propsPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
                    using (var stream = propsPart.GetStream(FileMode.Create))
                    using (var writer = new StreamWriter(stream))
                    {
                        writer.Write(kvp.Value.PropertiesXml);
                    }
                }
            }
            catch
            {
                // Skip custom XML parts that fail to restore
            }
        }

        // Restore Glossary Document Part (for Quick Parts, building blocks, document property fields)
        RestoreGlossaryDocumentPart();
    }

    /// <summary>
    /// Restores the Glossary Document Part (building blocks, Quick Parts)
    /// </summary>
    private void RestoreGlossaryDocumentPart()
    {
        if (string.IsNullOrEmpty(_packageData?.GlossaryDocumentXml))
            return;

        try
        {
            var glossaryPart = _mainPart!.AddNewPart<GlossaryDocumentPart>();
            var cleanedGlossaryXml = CleanXmlAttributes(_packageData.GlossaryDocumentXml);
            glossaryPart.GlossaryDocument = new GlossaryDocument(cleanedGlossaryXml);
            FixIndentationAttributes(glossaryPart.GlossaryDocument);

            // Restore glossary styles
            if (!string.IsNullOrEmpty(_packageData.GlossaryStylesXml))
            {
                var glossaryStylesPart = glossaryPart.AddNewPart<StyleDefinitionsPart>();
                var cleanedStylesXml = CleanXmlAttributes(_packageData.GlossaryStylesXml);
                glossaryStylesPart.Styles = new Styles(cleanedStylesXml);
                FixIndentationAttributes(glossaryStylesPart.Styles);
            }

            // Restore glossary font table
            if (!string.IsNullOrEmpty(_packageData.GlossaryFontTableXml))
            {
                var glossaryFontPart = glossaryPart.AddNewPart<FontTablePart>();
                var cleanedFontXml = CleanXmlAttributes(_packageData.GlossaryFontTableXml);
                glossaryFontPart.Fonts = new Fonts(cleanedFontXml);
            }

            // Restore glossary images
            var glossaryImageMapping = new Dictionary<string, string>();
            foreach (var kvp in _packageData.GlossaryImages)
            {
                var originalRelId = kvp.Key;
                var imageData = kvp.Value;

                var imagePart = glossaryPart.AddImagePart(imageData.ContentType);
                using (var stream = new MemoryStream(imageData.Data))
                {
                    imagePart.FeedData(stream);
                }

                var newRelId = glossaryPart.GetIdOfPart(imagePart);
                glossaryImageMapping[originalRelId] = newRelId;
            }

            // Update image references in glossary document
            if (glossaryImageMapping.Count > 0 && glossaryPart.GlossaryDocument != null)
            {
                var xml = glossaryPart.GlossaryDocument.OuterXml;
                foreach (var kvp in glossaryImageMapping)
                {
                    xml = xml.Replace($"r:embed=\"{kvp.Key}\"", $"r:embed=\"{kvp.Value}\"");
                    xml = xml.Replace($"r:id=\"{kvp.Key}\"", $"r:id=\"{kvp.Value}\"");
                }
                glossaryPart.GlossaryDocument = new GlossaryDocument(xml);
            }
        }
        catch
        {
            // Skip glossary document restoration if it fails
        }
    }

    /// <summary>
    /// Cleans problematic attributes from XML content without updating relationships
    /// </summary>
    private static string CleanXmlAttributes(string xml)
    {
        var result = xml;

        // Remove ALL w14: prefixed attributes (Word 2010 tracking attributes)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w14:[a-zA-Z0-9]+=""[^""]*""", "");

        // Remove w14: prefixed elements (self-closing first, then nested)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"<w14:[^>]*/\s*>", "");
        result = RemoveXmlElements(result, "w14:");

        // Remove wp14: prefixed attributes (Word 2010 drawing extensions)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+wp14:[a-zA-Z0-9]+=""[^""]*""", "");

        // Remove wp14: prefixed elements (Word 2010 drawing extensions)
        // Handle self-closing elements first
        result = System.Text.RegularExpressions.Regex.Replace(result, @"<wp14:[^>]*/\s*>", "");
        // Handle wp14 elements with nested content - need to handle nested wp14 elements properly
        result = RemoveXmlElements(result, "wp14:");

        // Remove w15/w16 prefixed attributes (Word 2013/2016 extensions)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w15:[a-zA-Z0-9]+=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w16[a-z]*:[a-zA-Z0-9]+=""[^""]*""", "");

        // Remove w15: prefixed elements
        result = System.Text.RegularExpressions.Regex.Replace(result, @"<w15:[^>]*/\s*>", "");
        result = RemoveXmlElements(result, "w15:");

        // Remove w16 prefixed elements (w16, w16se, w16cid, etc.)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"<w16[a-z]*:[^>]*/\s*>", "");
        result = RemoveXmlElements(result, "w16se:");
        result = RemoveXmlElements(result, "w16cid:");
        result = RemoveXmlElements(result, "w16cex:");
        result = RemoveXmlElements(result, "w16sdtdh:");
        result = RemoveXmlElements(result, "w16:");

        // Remove namespace declarations for removed prefixes
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+xmlns:w14=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+xmlns:wp14=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+xmlns:w15=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+xmlns:w16[a-z]*=""[^""]*""", "");

        // Remove rsid attributes (revision save IDs - not essential for display)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w:rsid[A-Za-z]*=""[^""]*""", "");

        // Remove w:start and w:end attributes entirely (Word 2010 RTL support - causes validation errors)
        // The SDK will fall back to w:left/w:right which are the standard attributes
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w:start=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w:end=""[^""]*""", "");

        // Clean mc:Ignorable attribute to remove references to namespaces we're stripping
        // This handles attributes like mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du"
        result = System.Text.RegularExpressions.Regex.Replace(result, @"mc:Ignorable=""[^""]*""", @"mc:Ignorable=""""");

        // Clean Requires attributes in mc:Choice elements to remove undefined prefixes
        // This is safer than trying to remove entire AlternateContent blocks
        result = System.Text.RegularExpressions.Regex.Replace(result, @"Requires=""wps""", @"Requires=""""");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"Requires=""wpc""", @"Requires=""""");

        return result;
    }

    /// <summary>
    /// Cleans header/footer XML with minimal modifications to preserve content.
    /// Only removes rsid attributes and updates mc:Ignorable to avoid validation issues,
    /// but keeps w14/w15/w16 attributes that are essential for preserving content structure.
    /// </summary>
    private static string CleanHeaderFooterXml(string xml)
    {
        var result = xml;

        // Only remove rsid attributes (revision save IDs) which can cause issues
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w:rsid[A-Za-z]*=""[^""]*""", "");

        return result;
    }

    /// <summary>
    /// Writes XML content directly to an OpenXML part, bypassing the strongly-typed API
    /// to preserve exact XML structure including extension elements.
    /// </summary>
    private static void WriteXmlToPart(OpenXmlPart part, string xml)
    {
        // Use UTF-8 encoding without BOM to match Word's format
        var encoding = new System.Text.UTF8Encoding(false);

        using (var stream = part.GetStream(FileMode.Create, FileAccess.Write))
        using (var writer = new StreamWriter(stream, encoding))
        {
            // Add XML declaration if not present
            if (!xml.TrimStart().StartsWith("<?xml", StringComparison.OrdinalIgnoreCase))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            }
            writer.Write(xml);
        }
    }

    /// <summary>
    /// Removes XML elements with the specified prefix, handling nested elements properly
    /// </summary>
    private static string RemoveXmlElements(string xml, string prefix)
    {
        var result = xml;
        bool changed;

        // Escape the prefix for regex (in case it contains special characters)
        var escapedPrefix = System.Text.RegularExpressions.Regex.Escape(prefix);

        // Keep iterating until no more elements with the prefix are found
        // This handles nested elements by removing innermost first
        do
        {
            changed = false;
            var openTagPattern = $@"<{escapedPrefix}([a-zA-Z0-9]+)(\s[^>]*)?>"; // Match opening tag
            var regex = new System.Text.RegularExpressions.Regex(openTagPattern);
            var match = regex.Match(result);

            while (match.Success)
            {
                var tagName = match.Groups[1].Value;
                var startIndex = match.Index;
                var closeTag = $"</{prefix}{tagName}>";
                var closeIndex = result.IndexOf(closeTag, startIndex + match.Length, StringComparison.Ordinal);

                if (closeIndex > 0)
                {
                    // Check if there's a nested element with the same prefix between open and close
                    var contentBetween = result.Substring(startIndex + match.Length, closeIndex - startIndex - match.Length);
                    var nestedMatch = System.Text.RegularExpressions.Regex.Match(contentBetween, $@"<{escapedPrefix}");

                    if (nestedMatch.Success)
                    {
                        // There's a nested element, skip this match and try the next one
                        match = match.NextMatch();
                        continue;
                    }

                    // Remove the entire element (from opening tag to closing tag inclusive)
                    var lengthToRemove = closeIndex + closeTag.Length - startIndex;
                    result = result.Remove(startIndex, lengthToRemove);
                    changed = true;
                    break; // Start over to handle any remaining elements
                }
                else
                {
                    // No closing tag found, try next match
                    match = match.NextMatch();
                }
            }
        } while (changed);

        return result;
    }

    /// <summary>
    /// Fixes Indentation elements that use Start/End instead of Left/Right
    /// </summary>
    private static void FixIndentationAttributes(OpenXmlElement element)
    {
        foreach (var ind in element.Descendants<Indentation>())
        {
            // Convert Start to Left if Start is set but Left is not
            if (ind.Start != null && ind.Left == null)
            {
                ind.Left = ind.Start.Value;
                ind.Start = null;
            }
            // Convert End to Right if End is set but Right is not
            if (ind.End != null && ind.Right == null)
            {
                ind.Right = ind.End.Value;
                ind.End = null;
            }
        }
    }

    /// <summary>
    /// Updates relationship IDs and cleans problematic attributes in XML content
    /// </summary>
    private string UpdateImageRelationships(string xml)
    {
        // First clean problematic attributes
        var result = CleanXmlAttributes(xml);

        // Update relationship IDs
        result = UpdateRelationshipIds(result);

        return result;
    }

    /// <summary>
    /// Updates relationship IDs in XML content with minimal cleaning.
    /// Removes only tracking attributes (paraId, textId, rsid) that cause validation errors
    /// but preserves all formatting elements for exact round-trip fidelity.
    /// </summary>
    private string UpdateRelationshipsOnly(string xml)
    {
        // Clean only tracking attributes that cause validation errors, preserve formatting
        var result = CleanTrackingAttributesOnly(xml);

        // Update relationship IDs
        return UpdateRelationshipIds(result);
    }

    /// <summary>
    /// Cleans only tracking/extension attributes and elements that cause validation errors.
    /// Removes Word 2010+ extension namespace content while preserving core formatting.
    /// </summary>
    private static string CleanTrackingAttributesOnly(string xml)
    {
        var result = xml;

        // Remove ALL w14: prefixed attributes and elements
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w14:[a-zA-Z0-9]+=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"<w14:[^>]*/\s*>", ""); // self-closing elements
        result = RemoveXmlElements(result, "w14:");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+xmlns:w14=""[^""]*""", "");

        // Remove wp14: prefixed attributes and elements (Word 2010 drawing extensions)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+wp14:[a-zA-Z0-9]+=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"<wp14:[^>]*/\s*>", "");
        result = RemoveXmlElements(result, "wp14:");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+xmlns:wp14=""[^""]*""", "");

        // Remove w15: prefixed attributes and elements (Word 2013 extensions)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w15:[a-zA-Z0-9]+=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"<w15:[^>]*/\s*>", "");
        result = RemoveXmlElements(result, "w15:");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+xmlns:w15=""[^""]*""", "");

        // Remove w16*: prefixed attributes and elements (Word 2016+ extensions)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w16[a-z]*:[a-zA-Z0-9]+=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"<w16[a-z]*:[^>]*/\s*>", "");
        result = RemoveXmlElements(result, "w16se:");
        result = RemoveXmlElements(result, "w16cid:");
        result = RemoveXmlElements(result, "w16cex:");
        result = RemoveXmlElements(result, "w16sdtdh:");
        result = RemoveXmlElements(result, "w16:");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+xmlns:w16[a-z]*=""[^""]*""", "");

        // Remove rsid attributes (revision save IDs - not needed for display)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w:rsid[A-Za-z]*=""[^""]*""", "");

        // Remove w:start and w:end attributes (Word 2010 RTL - causes validation errors)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w:start=""[^""]*""", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+w:end=""[^""]*""", "");

        return result;
    }

    /// <summary>
    /// Updates all relationship IDs in the XML string
    /// </summary>
    private string UpdateRelationshipIds(string xml)
    {
        var result = xml;

        // Update image relationship IDs
        foreach (var kvp in _imageRelationshipMapping)
        {
            result = result.Replace($"r:embed=\"{kvp.Key}\"", $"r:embed=\"{kvp.Value}\"");
            result = result.Replace($"r:id=\"{kvp.Key}\"", $"r:id=\"{kvp.Value}\"");
        }

        // Update hyperlink relationship IDs
        foreach (var kvp in _hyperlinkRelationshipMapping)
        {
            result = result.Replace($"r:id=\"{kvp.Key}\"", $"r:id=\"{kvp.Value}\"");
        }

        // Update header relationship IDs (for section properties in paragraphs)
        foreach (var kvp in _headerRelationshipMapping)
        {
            result = result.Replace($"r:id=\"{kvp.Key}\"", $"r:id=\"{kvp.Value}\"");
        }

        // Update footer relationship IDs (for section properties in paragraphs)
        foreach (var kvp in _footerRelationshipMapping)
        {
            result = result.Replace($"r:id=\"{kvp.Key}\"", $"r:id=\"{kvp.Value}\"");
        }

        return result;
    }

    /// <summary>
    /// Restores document properties
    /// </summary>
    private void RestoreDocumentProperties()
    {
        // Restore core properties
        if (_packageData?.CoreProperties != null)
        {
            var props = _document!.PackageProperties;
            var core = _packageData.CoreProperties;

            // Restore all text-based core properties (including empty strings for round-trip fidelity)
            if (core.Title != null) props.Title = core.Title;
            if (core.Subject != null) props.Subject = core.Subject;
            if (core.Creator != null) props.Creator = core.Creator;
            if (core.Keywords != null) props.Keywords = core.Keywords;
            if (core.Description != null) props.Description = core.Description;
            if (core.Category != null) props.Category = core.Category;
            if (core.LastModifiedBy != null) props.LastModifiedBy = core.LastModifiedBy;
            if (core.Revision != null) props.Revision = core.Revision;
            if (core.ContentStatus != null) props.ContentStatus = core.ContentStatus;

            // Restore date properties
            if (!string.IsNullOrEmpty(core.Created) && DateTime.TryParse(core.Created, out var created))
                props.Created = created;
            if (!string.IsNullOrEmpty(core.Modified) && DateTime.TryParse(core.Modified, out var modified))
                props.Modified = modified;
        }

        // Restore extended properties (Company, Template, Application, etc.)
        RestoreExtendedProperties();

        // Restore custom properties
        RestoreCustomProperties();
    }

    /// <summary>
    /// Restores extended document properties
    /// </summary>
    private void RestoreExtendedProperties()
    {
        if (_packageData?.ExtendedProperties == null)
            return;

        try
        {
            var extProps = _packageData.ExtendedProperties;

            // Get or create the extended properties part
            var extPropsPart = _document!.ExtendedFilePropertiesPart;
            if (extPropsPart == null)
            {
                extPropsPart = _document.AddExtendedFilePropertiesPart();
                extPropsPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
            }

            var properties = extPropsPart.Properties;

            // Restore text properties (including empty strings for round-trip fidelity)
            if (extProps.Template != null)
            {
                properties.Template = new DocumentFormat.OpenXml.ExtendedProperties.Template(extProps.Template);
            }
            if (extProps.Company != null)
            {
                properties.Company = new DocumentFormat.OpenXml.ExtendedProperties.Company(extProps.Company);
            }
            if (extProps.Manager != null)
            {
                properties.Manager = new DocumentFormat.OpenXml.ExtendedProperties.Manager(extProps.Manager);
            }
            if (extProps.Application != null)
            {
                properties.Application = new DocumentFormat.OpenXml.ExtendedProperties.Application(extProps.Application);
            }
            if (extProps.AppVersion != null)
            {
                properties.ApplicationVersion = new DocumentFormat.OpenXml.ExtendedProperties.ApplicationVersion(extProps.AppVersion);
            }

            // Restore numeric properties (these will be regenerated by Word on save, but include for completeness)
            if (extProps.Pages.HasValue)
            {
                properties.Pages = new DocumentFormat.OpenXml.ExtendedProperties.Pages(extProps.Pages.Value.ToString());
            }
            if (extProps.Words.HasValue)
            {
                properties.Words = new DocumentFormat.OpenXml.ExtendedProperties.Words(extProps.Words.Value.ToString());
            }
            if (extProps.Characters.HasValue)
            {
                properties.Characters = new DocumentFormat.OpenXml.ExtendedProperties.Characters(extProps.Characters.Value.ToString());
            }
            if (extProps.CharactersWithSpaces.HasValue)
            {
                properties.CharactersWithSpaces = new DocumentFormat.OpenXml.ExtendedProperties.CharactersWithSpaces(extProps.CharactersWithSpaces.Value.ToString());
            }
            if (extProps.Lines.HasValue)
            {
                properties.Lines = new DocumentFormat.OpenXml.ExtendedProperties.Lines(extProps.Lines.Value.ToString());
            }
            if (extProps.Paragraphs.HasValue)
            {
                properties.Paragraphs = new DocumentFormat.OpenXml.ExtendedProperties.Paragraphs(extProps.Paragraphs.Value.ToString());
            }
            if (extProps.TotalTime.HasValue)
            {
                properties.TotalTime = new DocumentFormat.OpenXml.ExtendedProperties.TotalTime(extProps.TotalTime.Value.ToString());
            }

            properties.Save();
        }
        catch
        {
            // Skip extended properties if restoration fails
        }
    }

    /// <summary>
    /// Restores custom document properties
    /// </summary>
    private void RestoreCustomProperties()
    {
        if (string.IsNullOrEmpty(_packageData?.CustomPropertiesXml))
            return;

        try
        {
            // Get or create the custom properties part
            var customPropsPart = _document!.CustomFilePropertiesPart;
            if (customPropsPart == null)
            {
                customPropsPart = _document.AddCustomFilePropertiesPart();
            }

            // Parse and restore the custom properties XML
            var cleanedXml = CleanXmlAttributes(_packageData.CustomPropertiesXml);
            customPropsPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(cleanedXml);
            customPropsPart.Properties.Save();
        }
        catch
        {
            // Skip custom properties if restoration fails
        }
    }

    /// <summary>
    /// Creates standard heading styles (Heading1 through Heading9)
    /// </summary>
    private void AddStyleDefinitions()
    {
        var stylesPart = _mainPart!.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();

        // Add default document style
        var defaultStyle = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        };
        defaultStyle.Append(new StyleName { Val = "Normal" });
        defaultStyle.Append(new PrimaryStyle());
        styles.Append(defaultStyle);

        // Add heading styles 1-9
        var headingSizes = new[] { 32, 26, 24, 22, 20, 18, 16, 15, 14 };
        var headingColors = new[] { "2F5496", "2F5496", "1F3763", "1F3763", "1F3763", "1F3763", "1F3763", "1F3763", "1F3763" };

        for (int i = 1; i <= 9; i++)
        {
            var headingStyle = CreateHeadingStyle(i, headingSizes[i - 1], headingColors[i - 1]);
            styles.Append(headingStyle);
        }

        // Add list paragraph style
        var listStyle = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "ListParagraph"
        };
        listStyle.Append(new StyleName { Val = "List Paragraph" });
        listStyle.Append(new BasedOn { Val = "Normal" });
        listStyle.Append(new StyleParagraphProperties(
            new Indentation { Left = "720" }
        ));
        styles.Append(listStyle);

        // Add hyperlink character style
        var hyperlinkStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = "Hyperlink"
        };
        hyperlinkStyle.Append(new StyleName { Val = "Hyperlink" });
        hyperlinkStyle.Append(new StyleRunProperties(
            new Color { Val = "0563C1" },
            new Underline { Val = UnderlineValues.Single }
        ));
        styles.Append(hyperlinkStyle);

        stylesPart.Styles = styles;
    }

    /// <summary>
    /// Creates a heading style definition
    /// </summary>
    private Style CreateHeadingStyle(int level, int fontSize, string color)
    {
        var style = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = $"Heading{level}"
        };

        style.Append(new StyleName { Val = $"heading {level}" });
        style.Append(new BasedOn { Val = "Normal" });
        style.Append(new NextParagraphStyle { Val = "Normal" });
        style.Append(new PrimaryStyle());

        var pPr = new StyleParagraphProperties();
        pPr.Append(new KeepNext());
        pPr.Append(new KeepLines());
        pPr.Append(new SpacingBetweenLines { Before = level == 1 ? "240" : "160", After = "80" });
        pPr.Append(new OutlineLevel { Val = level - 1 });
        style.Append(pPr);

        var rPr = new StyleRunProperties();
        rPr.Append(new Bold());
        rPr.Append(new FontSize { Val = (fontSize * 2).ToString() });
        rPr.Append(new FontSizeComplexScript { Val = (fontSize * 2).ToString() });
        rPr.Append(new Color { Val = color });
        style.Append(rPr);

        return style;
    }

    /// <summary>
    /// Processes a node and its children, writing content to the document
    /// </summary>
    private void ProcessNode(DocumentNode node)
    {
        switch (node.Type)
        {
            case ContentType.Document:
                foreach (var child in node.Children)
                {
                    ProcessNode(child);
                }
                break;

            case ContentType.Heading:
                WriteHeading(node);
                break;

            case ContentType.Paragraph:
                WriteParagraph(node);
                break;

            case ContentType.Table:
                WriteTable(node);
                break;

            case ContentType.Image:
                WriteImage(node);
                break;

            case ContentType.List:
                WriteList(node);
                break;

            case ContentType.ListItem:
                WriteListItem(node);
                break;

            case ContentType.HyperlinkText:
            case ContentType.TextRun:
                WriteParagraph(node);
                break;

            case ContentType.ContentControl:
                WriteSdtBlock(node);
                break;
        }
    }

    /// <summary>
    /// Writes a heading paragraph with full formatting
    /// </summary>
    private void WriteHeading(DocumentNode node)
    {
        // Check if this is an SDT block that should be preserved
        if (node.Metadata.TryGetValue("IsSdtBlock", out var isSdtBlock) && (bool)isSdtBlock)
        {
            WriteSdtBlock(node);
            return;
        }

        // Use original XML if available for exact round-trip
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            var updatedXml = UpdateImageRelationships(node.OriginalXml);

            // Check if the original XML is an SDT block
            if (updatedXml.TrimStart().StartsWith("<w:sdt"))
            {
                var sdtBlock = new SdtBlock(updatedXml);
                // Don't call UpdateSdtBlockContent - OriginalXml already has correct content
                // Modifying it could corrupt field codes and complex content
                FixIndentationAttributes(sdtBlock);
                _body!.Append(sdtBlock);
            }
            else
            {
                var paragraph = new Paragraph(updatedXml);
                // Don't call UpdateParagraphSdtContent - OriginalXml already has correct content
                FixIndentationAttributes(paragraph);
                _body!.Append(paragraph);
            }
        }
        else
        {
            var paragraph = new Paragraph();

            // Apply paragraph properties
            var paragraphProps = CreateParagraphProperties(node);
            if (paragraphProps.HasChildren)
            {
                paragraph.Append(paragraphProps);
            }

            // Ensure heading style is applied
            if (paragraph.ParagraphProperties == null)
            {
                paragraph.ParagraphProperties = new ParagraphProperties();
            }
            if (paragraph.ParagraphProperties.ParagraphStyleId == null)
            {
                paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = $"Heading{node.HeadingLevel}" };
            }

            // Write formatted runs or plain text
            WriteRunsOrText(paragraph, node);

            _body!.Append(paragraph);
        }

        // Process children (content under this heading)
        // Skip if we used original XML and it was an SDT block (content already included)
        bool skipChildren = !string.IsNullOrEmpty(node.OriginalXml) &&
                            (node.OriginalXml.TrimStart().StartsWith("<w:sdt") ||
                             node.Metadata.ContainsKey("IsSdtBlock"));

        if (!skipChildren)
        {
            foreach (var child in node.Children)
            {
                ProcessNode(child);
            }
        }
    }

    /// <summary>
    /// Writes a regular paragraph with full formatting
    /// </summary>
    private void WriteParagraph(DocumentNode node)
    {
        // Check if this is an SDT block that should be preserved
        if (node.Metadata.TryGetValue("IsSdtBlock", out var isSdtBlock) && (bool)isSdtBlock)
        {
            WriteSdtBlock(node);
            return;
        }

        // Use original XML if available for exact round-trip
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            var updatedXml = UpdateImageRelationships(node.OriginalXml);

            // Check if the original XML is an SDT block
            if (updatedXml.TrimStart().StartsWith("<w:sdt"))
            {
                var sdtBlock = new SdtBlock(updatedXml);
                // Don't call UpdateSdtBlockContent - OriginalXml already has correct content
                // Modifying it could corrupt field codes and complex content
                FixIndentationAttributes(sdtBlock);
                _body!.Append(sdtBlock);
            }
            else
            {
                var paragraph = new Paragraph(updatedXml);
                // Don't call UpdateParagraphSdtContent - OriginalXml already has correct content
                FixIndentationAttributes(paragraph);
                _body!.Append(paragraph);
            }
        }
        else
        {
            var paragraph = new Paragraph();

            // Apply paragraph properties
            var paragraphProps = CreateParagraphProperties(node);
            if (paragraphProps.HasChildren)
            {
                paragraph.Append(paragraphProps);
            }

            // Check for hyperlinks
            if (node.Metadata.TryGetValue("Hyperlinks", out var hyperlinksObj) && hyperlinksObj is List<HyperlinkData> hyperlinks)
            {
                WriteHyperlinkParagraph(paragraph, node, hyperlinks);
            }
            else
            {
                WriteRunsOrText(paragraph, node);
            }

            _body!.Append(paragraph);
        }

        // Process child images if any (skip if we used original XML as images are already included)
        if (string.IsNullOrEmpty(node.OriginalXml))
        {
            foreach (var child in node.Children)
            {
                if (child.Type == ContentType.Image)
                {
                    WriteImage(child);
                }
            }
        }

        // Process other children (skip if this was an SDT block)
        if (!node.Metadata.ContainsKey("IsSdtBlock") || !(bool)node.Metadata["IsSdtBlock"])
        {
            foreach (var child in node.Children)
            {
                if (child.Type != ContentType.Image)
                {
                    ProcessNode(child);
                }
            }
        }
    }

    /// <summary>
    /// Writes an SDT block (Structured Document Tag) with preserved original XML
    /// </summary>
    private void WriteSdtBlock(DocumentNode node)
    {
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            // Use UpdateRelationshipsOnly to preserve all formatting (including tables inside SDT blocks)
            var updatedXml = UpdateRelationshipsOnly(node.OriginalXml);
            var sdtBlock = new SdtBlock(updatedXml);

            // Only update SDT content for simple content controls that have been explicitly modified.
            // Do NOT update complex structures like TOC, tables, etc. - they should use OriginalXml exactly.
            // Check if this is a simple content control (single paragraph with content control properties)
            // and not a complex structure (multiple children, or types like TOC/Bibliography)
            var isSimpleContentControl = node.ContentControlProperties != null &&
                                          node.Children.Count <= 1 &&
                                          node.ContentControlProperties.Type != Models.ContentControls.ContentControlType.Unknown &&
                                          node.ContentControlProperties.Type != Models.ContentControls.ContentControlType.Group &&
                                          node.ContentControlProperties.Type != Models.ContentControls.ContentControlType.Bibliography &&
                                          !node.Metadata.ContainsKey("IsSdtBlock");

            if (isSimpleContentControl)
            {
                UpdateSdtBlockContent(sdtBlock, node);
            }

            FixIndentationAttributes(sdtBlock);
            _body!.Append(sdtBlock);
        }
        else
        {
            // Fallback: write children as regular content if original XML not available
            foreach (var child in node.Children)
            {
                ProcessNode(child);
            }
        }
    }

    /// <summary>
    /// Updates the content of an SDT block based on modified node values
    /// </summary>
    private void UpdateSdtBlockContent(SdtBlock sdtBlock, DocumentNode node)
    {
        var ccProps = node.ContentControlProperties;
        if (ccProps == null) return;

        // Get the SDT content element
        var sdtContent = sdtBlock.SdtContentBlock;
        if (sdtContent == null) return;

        // Get the new text value
        var newText = node.HasFormattedRuns
            ? string.Concat(node.Runs.Select(r => r.IsTab ? "\t" : r.IsBreak ? "" : r.Text))
            : node.Text;

        // Update the SDT properties based on control type
        var sdtPr = sdtBlock.SdtProperties;
        if (sdtPr != null)
        {
            UpdateSdtProperties(sdtPr, ccProps);
        }

        // Update the text content within the SDT
        // Find all Text elements within the content and update them
        var textElements = sdtContent.Descendants<Text>().ToList();
        if (textElements.Count > 0)
        {
            // Clear all but the first text element and update the first one
            var firstText = textElements[0];
            firstText.Text = newText;

            // Remove extra text elements (if any)
            for (int i = 1; i < textElements.Count; i++)
            {
                textElements[i].Remove();
            }
        }
        else
        {
            // No text elements found, try to add text to the first run in the first paragraph
            var firstPara = sdtContent.GetFirstChild<Paragraph>();
            if (firstPara != null)
            {
                var firstRun = firstPara.GetFirstChild<Run>();
                if (firstRun != null)
                {
                    // Remove any existing text children
                    foreach (var existingText in firstRun.Elements<Text>().ToList())
                    {
                        existingText.Remove();
                    }
                    firstRun.Append(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });
                }
            }
        }
    }

    /// <summary>
    /// Updates the SDT properties based on the content control properties
    /// </summary>
    private void UpdateSdtProperties(SdtProperties sdtPr, ContentControlProperties ccProps)
    {
        switch (ccProps.Type)
        {
            case ContentControlType.Checkbox:
                UpdateCheckboxProperties(sdtPr, ccProps);
                break;
            case ContentControlType.Date:
                UpdateDateProperties(sdtPr, ccProps);
                break;
            case ContentControlType.DropDownList:
            case ContentControlType.ComboBox:
                // For dropdowns and comboboxes, the text content is the main update
                // The selected value is represented by the displayed text
                break;
        }
    }

    /// <summary>
    /// Updates checkbox-specific properties in SDT
    /// </summary>
    private void UpdateCheckboxProperties(SdtProperties sdtPr, ContentControlProperties ccProps)
    {
        // Find the w14:checkbox element
        var checkbox = sdtPr.Descendants().FirstOrDefault(e => e.LocalName == "checkbox");
        if (checkbox != null)
        {
            // Find the w14:checked element
            var checkedElement = checkbox.Descendants().FirstOrDefault(e => e.LocalName == "checked");
            if (checkedElement != null)
            {
                // Update the val attribute
                var valAttr = checkedElement.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
                var newValue = ccProps.IsChecked == true ? "1" : "0";

                if (valAttr.LocalName != null)
                {
                    // Create a new OpenXmlAttribute with the updated value
                    var attrs = checkedElement.GetAttributes().ToList();
                    checkedElement.ClearAllAttributes();
                    foreach (var attr in attrs)
                    {
                        if (attr.LocalName == "val")
                        {
                            checkedElement.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newValue));
                        }
                        else
                        {
                            checkedElement.SetAttribute(attr);
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Updates date-specific properties in SDT
    /// </summary>
    private void UpdateDateProperties(SdtProperties sdtPr, ContentControlProperties ccProps)
    {
        var datePr = sdtPr.GetFirstChild<SdtContentDate>();
        if (datePr != null && ccProps.DateValue.HasValue)
        {
            datePr.FullDate = ccProps.DateValue.Value;
        }
    }

    /// <summary>
    /// Updates inline SDT content (SdtRun) within a paragraph based on modified node runs
    /// </summary>
    private void UpdateParagraphSdtContent(Paragraph paragraph, DocumentNode node)
    {
        // Find all SdtRun elements in the paragraph
        var sdtRuns = paragraph.Descendants<SdtRun>().ToList();
        if (sdtRuns.Count == 0 || node.Runs.Count == 0) return;

        // Build a map of content control IDs to their new values
        var ccValueMap = new Dictionary<int, (string text, ContentControlProperties props)>();
        foreach (var run in node.Runs)
        {
            if (run.ContentControlProperties != null && run.ContentControlProperties.Id.HasValue)
            {
                var id = run.ContentControlProperties.Id.Value;
                if (!ccValueMap.ContainsKey(id))
                {
                    ccValueMap[id] = (run.Text, run.ContentControlProperties);
                }
                else
                {
                    // Append text for runs with the same content control ID
                    var existing = ccValueMap[id];
                    ccValueMap[id] = (existing.text + run.Text, run.ContentControlProperties);
                }
            }
        }

        // Update each SdtRun
        foreach (var sdtRun in sdtRuns)
        {
            var sdtPr = sdtRun.SdtProperties;
            var sdtContent = sdtRun.SdtContentRun;

            if (sdtPr == null || sdtContent == null) continue;

            // Get the ID of this SDT
            var sdtId = sdtPr.GetFirstChild<SdtId>()?.Val?.Value;
            if (sdtId.HasValue && ccValueMap.TryGetValue(sdtId.Value, out var newValue))
            {
                // Update properties if needed (e.g., checkbox state)
                UpdateSdtProperties(sdtPr, newValue.props);

                // Update text content
                var textElements = sdtContent.Descendants<Text>().ToList();
                if (textElements.Count > 0)
                {
                    var firstText = textElements[0];
                    firstText.Text = newValue.text;

                    // Remove extra text elements
                    for (int i = 1; i < textElements.Count; i++)
                    {
                        textElements[i].Remove();
                    }
                }
                else
                {
                    // Try to add text to the first run
                    var firstRun = sdtContent.GetFirstChild<Run>();
                    if (firstRun != null)
                    {
                        foreach (var existingText in firstRun.Elements<Text>().ToList())
                        {
                            existingText.Remove();
                        }
                        firstRun.Append(new Text(newValue.text) { Space = SpaceProcessingModeValues.Preserve });
                    }
                }
            }
        }
    }

    /// <summary>
    /// Creates paragraph properties from formatting.
    /// Elements are added in the correct OOXML sequence order:
    /// pStyle, keepNext, keepLines, pageBreakBefore, widowControl, numPr, pBdr, shd, spacing, ind, jc
    /// </summary>
    private ParagraphProperties CreateParagraphProperties(DocumentNode node)
    {
        var props = new ParagraphProperties();
        var fmt = node.ParagraphFormatting;

        if (fmt == null) return props;

        // 1. Style (pStyle)
        if (!string.IsNullOrEmpty(fmt.StyleId))
        {
            props.Append(new ParagraphStyleId { Val = fmt.StyleId });
        }

        // 2-5. Keep with next/keep lines/page break/widow control
        if (fmt.KeepNext) props.Append(new KeepNext());
        if (fmt.KeepLines) props.Append(new KeepLines());
        if (fmt.PageBreakBefore) props.Append(new PageBreakBefore());
        if (fmt.WidowControl) props.Append(new WidowControl());

        // 7. Numbering (numPr) - must come before borders and shading
        if (fmt.NumberingId.HasValue)
        {
            props.Append(new NumberingProperties(
                new NumberingLevelReference { Val = fmt.NumberingLevel ?? 0 },
                new NumberingId { Val = fmt.NumberingId.Value }
            ));
        }

        // 9. Borders (pBdr) - must come before shading
        if (fmt.TopBorder != null || fmt.BottomBorder != null || fmt.LeftBorder != null || fmt.RightBorder != null)
        {
            var borders = new ParagraphBorders();
            // Border order within pBdr: top, left, bottom, right
            if (fmt.TopBorder != null) borders.Append(CreateBorder<TopBorder>(fmt.TopBorder));
            if (fmt.LeftBorder != null) borders.Append(CreateBorder<LeftBorder>(fmt.LeftBorder));
            if (fmt.BottomBorder != null) borders.Append(CreateBorder<BottomBorder>(fmt.BottomBorder));
            if (fmt.RightBorder != null) borders.Append(CreateBorder<RightBorder>(fmt.RightBorder));
            props.Append(borders);
        }

        // 10. Shading (shd)
        if (!string.IsNullOrEmpty(fmt.ShadingFill))
        {
            props.Append(new Shading { Fill = fmt.ShadingFill, Color = fmt.ShadingColor });
        }

        // 22. Spacing
        if (fmt.SpacingBefore != null || fmt.SpacingAfter != null || fmt.LineSpacing != null)
        {
            var spacing = new SpacingBetweenLines();
            if (fmt.SpacingBefore != null) spacing.Before = fmt.SpacingBefore;
            if (fmt.SpacingAfter != null) spacing.After = fmt.SpacingAfter;
            if (fmt.LineSpacing != null) spacing.Line = fmt.LineSpacing;
            if (!string.IsNullOrEmpty(fmt.LineSpacingRule))
            {
                spacing.LineRule = fmt.LineSpacingRule switch
                {
                    "Auto" => LineSpacingRuleValues.Auto,
                    "Exact" => LineSpacingRuleValues.Exact,
                    "AtLeast" => LineSpacingRuleValues.AtLeast,
                    _ => null
                };
            }
            props.Append(spacing);
        }

        // 23. Indentation (ind)
        if (fmt.IndentLeft != null || fmt.IndentRight != null || fmt.IndentFirstLine != null || fmt.IndentHanging != null)
        {
            var ind = new Indentation();
            if (fmt.IndentLeft != null) ind.Left = fmt.IndentLeft;
            if (fmt.IndentRight != null) ind.Right = fmt.IndentRight;
            if (fmt.IndentFirstLine != null) ind.FirstLine = fmt.IndentFirstLine;
            if (fmt.IndentHanging != null) ind.Hanging = fmt.IndentHanging;
            props.Append(ind);
        }

        // 27. Alignment (jc) - must come near the end
        if (!string.IsNullOrEmpty(fmt.Alignment))
        {
            var justification = fmt.Alignment switch
            {
                "Left" => JustificationValues.Left,
                "Center" => JustificationValues.Center,
                "Right" => JustificationValues.Right,
                "Both" => JustificationValues.Both,
                _ => (JustificationValues?)null
            };
            if (justification.HasValue)
            {
                props.Append(new Justification { Val = justification.Value });
            }
        }

        return props;
    }

    /// <summary>
    /// Creates a border element from formatting
    /// </summary>
    private T CreateBorder<T>(BorderFormatting fmt) where T : BorderType, new()
    {
        var border = new T();
        if (!string.IsNullOrEmpty(fmt.Style))
        {
            border.Val = fmt.Style switch
            {
                "Single" => BorderValues.Single,
                "Double" => BorderValues.Double,
                "Dashed" => BorderValues.Dashed,
                "Dotted" => BorderValues.Dotted,
                "Thick" => BorderValues.Thick,
                "None" => BorderValues.None,
                _ => BorderValues.Single
            };
        }
        if (!string.IsNullOrEmpty(fmt.Size) && uint.TryParse(fmt.Size, out var size))
        {
            border.Size = size;
        }
        if (!string.IsNullOrEmpty(fmt.Color))
        {
            border.Color = fmt.Color;
        }
        if (!string.IsNullOrEmpty(fmt.Space) && uint.TryParse(fmt.Space, out var space))
        {
            border.Space = space;
        }
        return border;
    }

    /// <summary>
    /// Writes runs with formatting or plain text
    /// </summary>
    private void WriteRunsOrText(Paragraph paragraph, DocumentNode node)
    {
        if (node.HasFormattedRuns)
        {
            foreach (var formattedRun in node.Runs)
            {
                var run = CreateRun(formattedRun);
                paragraph.Append(run);
            }
        }
        else if (!string.IsNullOrEmpty(node.Text))
        {
            var run = new Run();
            run.Append(new Text(node.Text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);
        }
    }

    /// <summary>
    /// Creates a run element from a formatted run
    /// </summary>
    private Run CreateRun(FormattedRun formattedRun)
    {
        var run = new Run();

        // Add run properties if there's formatting
        var runProps = CreateRunProperties(formattedRun.Formatting);
        if (runProps.HasChildren)
        {
            run.Append(runProps);
        }

        // Add content
        if (formattedRun.IsTab)
        {
            run.Append(new TabChar());
        }
        else if (formattedRun.IsBreak)
        {
            var br = new Break();
            if (formattedRun.BreakType == "Page")
                br.Type = BreakValues.Page;
            else if (formattedRun.BreakType == "Column")
                br.Type = BreakValues.Column;
            run.Append(br);
        }
        else if (!string.IsNullOrEmpty(formattedRun.Text))
        {
            run.Append(new Text(formattedRun.Text) { Space = SpaceProcessingModeValues.Preserve });
        }

        return run;
    }

    /// <summary>
    /// Creates run properties from formatting.
    /// Elements are added in the correct OOXML sequence order:
    /// rStyle, rFonts, b, i, caps, smallCaps, strike, dstrike, color, sz, szCs, highlight, u, shd, vertAlign
    /// </summary>
    private RunProperties CreateRunProperties(RunFormatting fmt)
    {
        var props = new RunProperties();

        // 1. Style (rStyle)
        if (!string.IsNullOrEmpty(fmt.StyleId))
        {
            props.Append(new RunStyle { Val = fmt.StyleId });
        }

        // 2. Font (rFonts) - must come early, after rStyle
        if (fmt.FontFamily != null || fmt.FontFamilyAscii != null)
        {
            var fonts = new RunFonts();
            if (fmt.FontFamily != null) fonts.HighAnsi = fmt.FontFamily;
            if (fmt.FontFamilyAscii != null) fonts.Ascii = fmt.FontFamilyAscii;
            if (fmt.FontFamilyEastAsia != null) fonts.EastAsia = fmt.FontFamilyEastAsia;
            if (fmt.FontFamilyComplexScript != null) fonts.ComplexScript = fmt.FontFamilyComplexScript;
            props.Append(fonts);
        }

        // 3. Bold (b)
        if (fmt.Bold)
        {
            props.Append(new Bold());
        }

        // 4. Italic (i)
        if (fmt.Italic)
        {
            props.Append(new Italic());
        }

        // 5. Caps
        if (fmt.AllCaps)
        {
            props.Append(new Caps());
        }

        // 6. SmallCaps
        if (fmt.SmallCaps)
        {
            props.Append(new SmallCaps());
        }

        // 7. Strike
        if (fmt.Strike)
        {
            props.Append(new Strike());
        }

        // 8. DoubleStrike (dstrike)
        if (fmt.DoubleStrike)
        {
            props.Append(new DoubleStrike());
        }

        // 9. Color
        if (!string.IsNullOrEmpty(fmt.Color))
        {
            props.Append(new Color { Val = fmt.Color });
        }

        // 10. Font size (sz)
        if (!string.IsNullOrEmpty(fmt.FontSize))
        {
            props.Append(new FontSize { Val = fmt.FontSize });
        }

        // 11. Font size complex script (szCs)
        if (!string.IsNullOrEmpty(fmt.FontSizeComplexScript))
        {
            props.Append(new FontSizeComplexScript { Val = fmt.FontSizeComplexScript });
        }

        // 12. Highlight
        if (!string.IsNullOrEmpty(fmt.Highlight))
        {
            var highlightValue = fmt.Highlight switch
            {
                "Yellow" => HighlightColorValues.Yellow,
                "Green" => HighlightColorValues.Green,
                "Cyan" => HighlightColorValues.Cyan,
                "Magenta" => HighlightColorValues.Magenta,
                "Blue" => HighlightColorValues.Blue,
                "Red" => HighlightColorValues.Red,
                "DarkBlue" => HighlightColorValues.DarkBlue,
                "DarkCyan" => HighlightColorValues.DarkCyan,
                "DarkGreen" => HighlightColorValues.DarkGreen,
                "DarkMagenta" => HighlightColorValues.DarkMagenta,
                "DarkRed" => HighlightColorValues.DarkRed,
                "DarkYellow" => HighlightColorValues.DarkYellow,
                "DarkGray" => HighlightColorValues.DarkGray,
                "LightGray" => HighlightColorValues.LightGray,
                "Black" => HighlightColorValues.Black,
                _ => (HighlightColorValues?)null
            };
            if (highlightValue.HasValue)
            {
                props.Append(new Highlight { Val = highlightValue.Value });
            }
        }

        // 13. Underline (u)
        if (fmt.Underline)
        {
            var underline = new Underline();
            if (!string.IsNullOrEmpty(fmt.UnderlineStyle))
            {
                underline.Val = fmt.UnderlineStyle switch
                {
                    "Single" => UnderlineValues.Single,
                    "Double" => UnderlineValues.Double,
                    "Wave" => UnderlineValues.Wave,
                    "Dotted" => UnderlineValues.Dotted,
                    "Dash" => UnderlineValues.Dash,
                    _ => UnderlineValues.Single
                };
            }
            else
            {
                underline.Val = UnderlineValues.Single;
            }
            props.Append(underline);
        }

        // 14. Shading (shd)
        if (!string.IsNullOrEmpty(fmt.Shading))
        {
            props.Append(new Shading { Fill = fmt.Shading });
        }

        // 15. Superscript/Subscript (vertAlign) - must come near the end
        if (fmt.Superscript)
        {
            props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        }
        else if (fmt.Subscript)
        {
            props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
        }

        return props;
    }

    /// <summary>
    /// Writes a paragraph containing hyperlinks
    /// </summary>
    private void WriteHyperlinkParagraph(Paragraph paragraph, DocumentNode node, List<HyperlinkData> hyperlinks)
    {
        // For now, write formatted runs with hyperlink style
        // Full hyperlink reconstruction requires more complex handling
        if (node.HasFormattedRuns)
        {
            foreach (var formattedRun in node.Runs)
            {
                var run = CreateRun(formattedRun);
                paragraph.Append(run);
            }
        }
        else if (!string.IsNullOrEmpty(node.Text))
        {
            // Try to create actual hyperlinks
            int lastIndex = 0;
            foreach (var linkData in hyperlinks)
            {
                if (!string.IsNullOrEmpty(linkData.Url) && !string.IsNullOrEmpty(linkData.Text))
                {
                    // Find the hyperlink text in the paragraph text
                    int linkIndex = node.Text.IndexOf(linkData.Text, lastIndex, StringComparison.Ordinal);
                    if (linkIndex >= 0)
                    {
                        // Write text before the link
                        if (linkIndex > lastIndex)
                        {
                            var beforeText = node.Text.Substring(lastIndex, linkIndex - lastIndex);
                            paragraph.Append(new Run(new Text(beforeText) { Space = SpaceProcessingModeValues.Preserve }));
                        }

                        // Create hyperlink relationship
                        var relId = GetOrCreateHyperlinkRelationship(linkData.Url);

                        // Create hyperlink element
                        var hyperlink = new Hyperlink { Id = relId };
                        if (!string.IsNullOrEmpty(linkData.Tooltip))
                        {
                            hyperlink.Tooltip = linkData.Tooltip;
                        }

                        // Add runs to hyperlink
                        if (linkData.Runs.Count > 0)
                        {
                            foreach (var formattedRun in linkData.Runs)
                            {
                                hyperlink.Append(CreateRun(formattedRun));
                            }
                        }
                        else
                        {
                            var linkRun = new Run();
                            linkRun.Append(new RunProperties(
                                new Color { Val = "0563C1" },
                                new Underline { Val = UnderlineValues.Single }
                            ));
                            linkRun.Append(new Text(linkData.Text) { Space = SpaceProcessingModeValues.Preserve });
                            hyperlink.Append(linkRun);
                        }

                        paragraph.Append(hyperlink);
                        lastIndex = linkIndex + linkData.Text.Length;
                    }
                }
            }

            // Write remaining text
            if (lastIndex < node.Text.Length)
            {
                var remainingText = node.Text.Substring(lastIndex);
                paragraph.Append(new Run(new Text(remainingText) { Space = SpaceProcessingModeValues.Preserve }));
            }
        }
    }

    /// <summary>
    /// Gets or creates a hyperlink relationship for a URL
    /// </summary>
    private string GetOrCreateHyperlinkRelationship(string url)
    {
        if (_hyperlinkRelationships.TryGetValue(url, out var existingId))
        {
            return existingId;
        }

        var uri = new Uri(url, UriKind.RelativeOrAbsolute);
        var rel = _mainPart!.AddHyperlinkRelationship(uri, true);
        _hyperlinkRelationships[url] = rel.Id;
        return rel.Id;
    }

    /// <summary>
    /// Writes a table to the document with full formatting
    /// </summary>
    private void WriteTable(DocumentNode node)
    {
        // Use original XML if available for exact round-trip
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            // Use UpdateRelationshipsOnly to preserve all table formatting (w14/w15/w16 elements)
            var updatedXml = UpdateRelationshipsOnly(node.OriginalXml);
            var originalTable = new Table(updatedXml);

            // Apply modifications from TableData to the parsed table
            var modifiedTableData = node.GetTableData();
            if (modifiedTableData != null)
            {
                ApplyTableDataModifications(originalTable, modifiedTableData);
            }

            FixIndentationAttributes(originalTable);
            _body!.Append(originalTable);
            // Don't add extra paragraph - preserve original document structure
            return;
        }

        var tableData = node.GetTableData();
        if (tableData == null || tableData.Rows.Count == 0)
        {
            var placeholder = new Paragraph();
            placeholder.Append(new Run(new Text("[Table]")));
            _body!.Append(placeholder);
            return;
        }

        var table = new Table();

        // Table properties with formatting
        var tableProps = CreateTableProperties(tableData.Formatting);
        table.Append(tableProps);

        // Table grid
        var tableGrid = new TableGrid();
        if (tableData.Formatting?.GridColumnWidths != null)
        {
            foreach (var width in tableData.Formatting.GridColumnWidths)
            {
                var col = new GridColumn();
                if (!string.IsNullOrEmpty(width))
                {
                    col.Width = width;
                }
                tableGrid.Append(col);
            }
        }
        else
        {
            for (int i = 0; i < tableData.ColumnCount; i++)
            {
                tableGrid.Append(new GridColumn());
            }
        }
        table.Append(tableGrid);

        // Table rows
        foreach (var rowData in tableData.Rows)
        {
            var tableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

            // Row properties
            var rowProps = CreateTableRowProperties(rowData.Formatting);
            if (rowProps.HasChildren)
            {
                tableRow.Append(rowProps);
            }

            foreach (var cellData in rowData.Cells)
            {
                var tableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();

                // Cell properties
                var cellProps = CreateTableCellProperties(cellData.Formatting);
                tableCell.Append(cellProps);

                // Cell content
                if (cellData.Content.Count > 0)
                {
                    foreach (var contentNode in cellData.Content)
                    {
                        var cellPara = new Paragraph();

                        // Apply paragraph formatting if available
                        if (contentNode.ParagraphFormatting != null)
                        {
                            var paraProps = CreateParagraphProperties(contentNode);
                            if (paraProps.HasChildren)
                            {
                                cellPara.Append(paraProps);
                            }
                        }

                        // Write formatted runs or text
                        if (contentNode.HasFormattedRuns)
                        {
                            foreach (var formattedRun in contentNode.Runs)
                            {
                                cellPara.Append(CreateRun(formattedRun));
                            }
                        }
                        else
                        {
                            cellPara.Append(new Run(new Text(contentNode.Text) { Space = SpaceProcessingModeValues.Preserve }));
                        }

                        tableCell.Append(cellPara);
                    }
                }
                else
                {
                    tableCell.Append(new Paragraph());
                }

                tableRow.Append(tableCell);
            }

            table.Append(tableRow);
        }

        _body!.Append(table);
        _body.Append(new Paragraph());
    }

    /// <summary>
    /// Creates table properties from formatting
    /// </summary>
    private TableProperties CreateTableProperties(TableFormatting? fmt)
    {
        var props = new TableProperties();

        if (fmt != null)
        {
            // Width
            if (!string.IsNullOrEmpty(fmt.Width))
            {
                var widthType = fmt.WidthType switch
                {
                    "Pct" => TableWidthUnitValues.Pct,
                    "Dxa" => TableWidthUnitValues.Dxa,
                    "Auto" => TableWidthUnitValues.Auto,
                    _ => TableWidthUnitValues.Dxa
                };
                props.Append(new TableWidth { Width = fmt.Width, Type = widthType });
            }
            else
            {
                props.Append(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct });
            }

            // Alignment
            if (!string.IsNullOrEmpty(fmt.Alignment))
            {
                var justification = fmt.Alignment switch
                {
                    "Left" => TableRowAlignmentValues.Left,
                    "Center" => TableRowAlignmentValues.Center,
                    "Right" => TableRowAlignmentValues.Right,
                    _ => (TableRowAlignmentValues?)null
                };
                if (justification.HasValue)
                {
                    props.Append(new TableJustification { Val = justification.Value });
                }
            }

            // Borders
            var borders = new TableBorders();
            bool hasBorders = false;

            if (fmt.TopBorder != null)
            {
                borders.Append(CreateBorder<TopBorder>(fmt.TopBorder));
                hasBorders = true;
            }
            if (fmt.BottomBorder != null)
            {
                borders.Append(CreateBorder<BottomBorder>(fmt.BottomBorder));
                hasBorders = true;
            }
            if (fmt.LeftBorder != null)
            {
                borders.Append(CreateBorder<LeftBorder>(fmt.LeftBorder));
                hasBorders = true;
            }
            if (fmt.RightBorder != null)
            {
                borders.Append(CreateBorder<RightBorder>(fmt.RightBorder));
                hasBorders = true;
            }
            if (fmt.InsideHorizontalBorder != null)
            {
                borders.Append(CreateBorder<InsideHorizontalBorder>(fmt.InsideHorizontalBorder));
                hasBorders = true;
            }
            if (fmt.InsideVerticalBorder != null)
            {
                borders.Append(CreateBorder<InsideVerticalBorder>(fmt.InsideVerticalBorder));
                hasBorders = true;
            }

            if (hasBorders)
            {
                props.Append(borders);
            }
            else
            {
                // Default borders
                props.Append(new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                ));
            }
        }
        else
        {
            props.Append(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct });
            props.Append(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
            ));
        }

        return props;
    }

    /// <summary>
    /// Creates table row properties from formatting
    /// </summary>
    private TableRowProperties CreateTableRowProperties(TableRowFormatting? fmt)
    {
        var props = new TableRowProperties();

        if (fmt != null)
        {
            // Height
            if (!string.IsNullOrEmpty(fmt.Height) && uint.TryParse(fmt.Height, out var height))
            {
                var heightRule = fmt.HeightRule switch
                {
                    "Exact" => HeightRuleValues.Exact,
                    "AtLeast" => HeightRuleValues.AtLeast,
                    _ => HeightRuleValues.Auto
                };
                props.Append(new TableRowHeight { Val = height, HeightType = heightRule });
            }

            // Header
            if (fmt.IsHeader)
            {
                props.Append(new TableHeader());
            }

            // Can't split
            if (fmt.CantSplit)
            {
                props.Append(new CantSplit());
            }
        }

        return props;
    }

    /// <summary>
    /// Creates table cell properties from formatting.
    /// Elements are added in the correct OOXML sequence order:
    /// tcW, gridSpan, vMerge, tcBorders, shd, noWrap, vAlign
    /// </summary>
    private TableCellProperties CreateTableCellProperties(TableCellFormatting? fmt)
    {
        var props = new TableCellProperties();

        if (fmt != null)
        {
            // 1. Width (tcW)
            if (!string.IsNullOrEmpty(fmt.Width))
            {
                var widthType = fmt.WidthType switch
                {
                    "Pct" => TableWidthUnitValues.Pct,
                    "Dxa" => TableWidthUnitValues.Dxa,
                    "Auto" => TableWidthUnitValues.Auto,
                    _ => TableWidthUnitValues.Dxa
                };
                props.Append(new TableCellWidth { Width = fmt.Width, Type = widthType });
            }

            // 2. Grid span
            if (fmt.GridSpan > 1)
            {
                props.Append(new GridSpan { Val = fmt.GridSpan });
            }

            // 3. Vertical merge (vMerge)
            if (!string.IsNullOrEmpty(fmt.VerticalMerge))
            {
                var vMerge = new VerticalMerge();
                if (fmt.VerticalMerge == "Restart")
                {
                    vMerge.Val = MergedCellValues.Restart;
                }
                props.Append(vMerge);
            }

            // 4. Borders (tcBorders) - must come before shading
            if (fmt.TopBorder != null || fmt.BottomBorder != null || fmt.LeftBorder != null || fmt.RightBorder != null)
            {
                var borders = new TableCellBorders();
                if (fmt.TopBorder != null) borders.Append(CreateBorder<TopBorder>(fmt.TopBorder));
                if (fmt.BottomBorder != null) borders.Append(CreateBorder<BottomBorder>(fmt.BottomBorder));
                if (fmt.LeftBorder != null) borders.Append(CreateBorder<LeftBorder>(fmt.LeftBorder));
                if (fmt.RightBorder != null) borders.Append(CreateBorder<RightBorder>(fmt.RightBorder));
                props.Append(borders);
            }

            // 5. Shading (shd) - must come after borders, before noWrap
            if (!string.IsNullOrEmpty(fmt.ShadingFill))
            {
                var shading = new Shading { Fill = fmt.ShadingFill };
                if (!string.IsNullOrEmpty(fmt.ShadingColor))
                {
                    shading.Color = fmt.ShadingColor;
                }
                if (!string.IsNullOrEmpty(fmt.ShadingPattern))
                {
                    shading.Val = fmt.ShadingPattern switch
                    {
                        "Clear" => ShadingPatternValues.Clear,
                        "Solid" => ShadingPatternValues.Solid,
                        _ => ShadingPatternValues.Clear
                    };
                }
                props.Append(shading);
            }

            // 6. No wrap - must come before vAlign
            if (fmt.NoWrap)
            {
                props.Append(new NoWrap());
            }

            // 7. Vertical alignment (vAlign) - must come near the end
            if (!string.IsNullOrEmpty(fmt.VerticalAlignment))
            {
                var vAlign = fmt.VerticalAlignment switch
                {
                    "Top" => TableVerticalAlignmentValues.Top,
                    "Center" => TableVerticalAlignmentValues.Center,
                    "Bottom" => TableVerticalAlignmentValues.Bottom,
                    _ => (TableVerticalAlignmentValues?)null
                };
                if (vAlign.HasValue)
                {
                    props.Append(new TableCellVerticalAlignment { Val = vAlign.Value });
                }
            }
        }

        return props;
    }

    /// <summary>
    /// Applies modifications from TableData to a parsed table XML element.
    /// Updates cell text, cell formatting (shading, borders), and row formatting.
    /// </summary>
    private void ApplyTableDataModifications(Table table, Models.Tables.TableData tableData)
    {
        var xmlRows = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ToList();

        for (var rowIndex = 0; rowIndex < tableData.Rows.Count && rowIndex < xmlRows.Count; rowIndex++)
        {
            var dataRow = tableData.Rows[rowIndex];
            var xmlRow = xmlRows[rowIndex];

            // Apply row-level formatting changes
            ApplyRowFormatting(xmlRow, dataRow);

            var xmlCells = xmlRow.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList();

            for (var cellIndex = 0; cellIndex < dataRow.Cells.Count && cellIndex < xmlCells.Count; cellIndex++)
            {
                var dataCell = dataRow.Cells[cellIndex];
                var xmlCell = xmlCells[cellIndex];

                // Apply cell-level formatting changes
                ApplyCellFormatting(xmlCell, dataCell);

                // Apply text changes to cell content
                ApplyCellTextChanges(xmlCell, dataCell);
            }
        }
    }

    /// <summary>
    /// Applies row-level formatting modifications to an XML table row.
    /// </summary>
    private void ApplyRowFormatting(DocumentFormat.OpenXml.Wordprocessing.TableRow xmlRow, Models.Tables.TableRow dataRow)
    {
        if (dataRow.Formatting == null) return;

        var rowProps = xmlRow.GetFirstChild<TableRowProperties>();
        if (rowProps == null)
        {
            rowProps = new TableRowProperties();
            xmlRow.InsertAt(rowProps, 0);
        }

        // Update header status
        var existingHeader = rowProps.GetFirstChild<TableHeader>();
        if (dataRow.Formatting.IsHeader && existingHeader == null)
        {
            rowProps.Append(new TableHeader());
        }
        else if (!dataRow.Formatting.IsHeader && existingHeader != null)
        {
            existingHeader.Remove();
        }
    }

    /// <summary>
    /// Applies cell-level formatting modifications to an XML table cell.
    /// Note: Borders are NOT modified here to preserve original XML element ordering.
    /// Border modifications should be done through explicit extension methods that set BordersModified flag.
    /// </summary>
    private void ApplyCellFormatting(DocumentFormat.OpenXml.Wordprocessing.TableCell xmlCell, Models.Tables.TableCell dataCell)
    {
        if (dataCell.Formatting == null) return;

        // NOTE: We intentionally do NOT process borders here.
        // The original XML already contains correctly-ordered border elements.
        // Modifying them would require careful handling of OOXML element ordering
        // (top, left/start, bottom, right/end, insideH, insideV) which is complex
        // and can cause validation errors. For now, borders are preserved as-is
        // from the original table XML.

        // Only process shading and vertical alignment which don't have ordering issues

        var cellProps = xmlCell.GetFirstChild<TableCellProperties>();
        if (cellProps == null)
        {
            // If there's no cell properties and nothing to modify, just return
            if (string.IsNullOrEmpty(dataCell.Formatting.ShadingFill) &&
                string.IsNullOrEmpty(dataCell.Formatting.VerticalAlignment))
            {
                return;
            }
            cellProps = new TableCellProperties();
            xmlCell.InsertAt(cellProps, 0);
        }

        // Update shading (shd) - only if explicitly set
        if (!string.IsNullOrEmpty(dataCell.Formatting.ShadingFill))
        {
            var existingShading = cellProps.GetFirstChild<Shading>();
            if (existingShading != null)
            {
                existingShading.Fill = dataCell.Formatting.ShadingFill;
            }
            else
            {
                var newShading = new Shading
                {
                    Fill = dataCell.Formatting.ShadingFill,
                    Val = ShadingPatternValues.Clear
                };
                InsertTableCellPropertyInOrder(cellProps, newShading);
            }
        }

        // Update vertical alignment (vAlign) - must come near the end
        if (!string.IsNullOrEmpty(dataCell.Formatting.VerticalAlignment))
        {
            var vAlign = dataCell.Formatting.VerticalAlignment.ToLowerInvariant() switch
            {
                "top" => TableVerticalAlignmentValues.Top,
                "center" => TableVerticalAlignmentValues.Center,
                "bottom" => TableVerticalAlignmentValues.Bottom,
                _ => (TableVerticalAlignmentValues?)null
            };
            if (vAlign.HasValue)
            {
                var existingAlign = cellProps.GetFirstChild<TableCellVerticalAlignment>();
                if (existingAlign != null)
                {
                    existingAlign.Val = vAlign.Value;
                }
                else
                {
                    var newAlign = new TableCellVerticalAlignment { Val = vAlign.Value };
                    InsertTableCellPropertyInOrder(cellProps, newAlign);
                }
            }
        }
    }

    /// <summary>
    /// Inserts a child element into TableCellProperties at the correct OOXML sequence position.
    /// OOXML order: cnfStyle, tcW, gridSpan, hMerge, vMerge, tcBorders, shd, noWrap, tcMar, textDirection, tcFitText, vAlign, hideMark
    /// </summary>
    private static void InsertTableCellPropertyInOrder(TableCellProperties cellProps, OpenXmlElement newElement)
    {
        // Define the correct order of tcPr child elements
        var elementOrder = new[]
        {
            typeof(ConditionalFormatStyle),     // cnfStyle
            typeof(TableCellWidth),              // tcW
            typeof(GridSpan),                    // gridSpan
            typeof(HorizontalMerge),             // hMerge
            typeof(VerticalMerge),               // vMerge
            typeof(TableCellBorders),            // tcBorders
            typeof(Shading),                     // shd
            typeof(NoWrap),                      // noWrap
            typeof(TableCellMargin),             // tcMar
            typeof(TextDirection),               // textDirection
            typeof(TableCellFitText),            // tcFitText
            typeof(TableCellVerticalAlignment),  // vAlign
            typeof(HideMark)                     // hideMark
        };

        var newElementIndex = Array.IndexOf(elementOrder, newElement.GetType());
        if (newElementIndex < 0)
        {
            // Unknown element type, append at the end
            cellProps.Append(newElement);
            return;
        }

        // Find the first element that should come after the new element
        foreach (var child in cellProps.ChildElements)
        {
            var childIndex = Array.IndexOf(elementOrder, child.GetType());
            if (childIndex > newElementIndex)
            {
                // Insert before this child
                child.InsertBeforeSelf(newElement);
                return;
            }
        }

        // No element found that should come after, so append at the end
        cellProps.Append(newElement);
    }

    /// <summary>
    /// Applies text changes from cell content nodes to the XML cell.
    /// Also handles nested tables recursively and font formatting.
    /// </summary>
    private void ApplyCellTextChanges(DocumentFormat.OpenXml.Wordprocessing.TableCell xmlCell, Models.Tables.TableCell dataCell)
    {
        if (dataCell.Content.Count == 0) return;

        var xmlParagraphs = xmlCell.Elements<Paragraph>().ToList();
        var xmlNestedTables = xmlCell.Elements<Table>().ToList();

        var paraIndex = 0;
        var tableIndex = 0;

        foreach (var contentNode in dataCell.Content)
        {
            // Handle nested tables - apply modifications recursively
            if (contentNode.Type == Core.ContentType.Table)
            {
                if (tableIndex < xmlNestedTables.Count)
                {
                    var nestedTableData = contentNode.GetTableData();
                    if (nestedTableData != null)
                    {
                        ApplyTableDataModifications(xmlNestedTables[tableIndex], nestedTableData);
                    }
                }
                tableIndex++;
                continue;
            }

            if (paraIndex < xmlParagraphs.Count)
            {
                var xmlPara = xmlParagraphs[paraIndex];

                // If the content node has formatted runs with font changes, update existing runs
                if (contentNode.HasFormattedRuns)
                {
                    // Update font properties on existing runs instead of rebuilding
                    // This preserves the original XML structure (bookmarks, fields, etc.)
                    ApplyFontToExistingRuns(xmlPara, contentNode.Runs);
                }
                else
                {
                    // Handle simple text changes (no formatting)
                    var textElements = xmlPara.Descendants<Text>().ToList();
                    if (textElements.Count > 0 && !string.IsNullOrEmpty(contentNode.Text))
                    {
                        var firstText = textElements[0];
                        var originalCombinedText = string.Join("", textElements.Select(t => t.Text));

                        if (contentNode.Text != originalCombinedText)
                        {
                            firstText.Text = contentNode.Text;
                            firstText.Space = SpaceProcessingModeValues.Preserve;

                            for (var i = 1; i < textElements.Count; i++)
                            {
                                textElements[i].Text = "";
                            }
                        }
                    }
                    else if (textElements.Count == 0 && !string.IsNullOrEmpty(contentNode.Text))
                    {
                        var run = xmlPara.GetFirstChild<Run>();
                        if (run == null)
                        {
                            run = new Run();
                            xmlPara.Append(run);
                        }
                        run.Append(new Text(contentNode.Text) { Space = SpaceProcessingModeValues.Preserve });
                    }
                }
            }
            paraIndex++;
        }
    }

    /// <summary>
    /// Applies font formatting from formatted runs to existing XML runs.
    /// This preserves the original XML structure while updating font properties.
    /// </summary>
    private void ApplyFontToExistingRuns(Paragraph xmlPara, List<Models.Formatting.FormattedRun> formattedRuns)
    {
        var xmlRuns = xmlPara.Elements<Run>().ToList();

        // If there are no XML runs but we have formatted runs, we need to create them
        if (xmlRuns.Count == 0 && formattedRuns.Count > 0)
        {
            foreach (var formattedRun in formattedRuns)
            {
                xmlPara.Append(CreateRun(formattedRun));
            }
            return;
        }

        // Get the primary font from formatted runs (use the first run with a font set)
        string? primaryFont = null;
        foreach (var frun in formattedRuns)
        {
            var font = frun.Formatting?.FontFamilyAscii ?? frun.Formatting?.FontFamily;
            if (!string.IsNullOrEmpty(font))
            {
                primaryFont = font;
                break;
            }
        }

        if (string.IsNullOrEmpty(primaryFont))
            return;

        // Apply font to all existing runs
        foreach (var xmlRun in xmlRuns)
        {
            var runProps = xmlRun.GetFirstChild<RunProperties>();
            if (runProps == null)
            {
                runProps = new RunProperties();
                xmlRun.InsertAt(runProps, 0);
            }

            // Update or add RunFonts
            var existingFonts = runProps.GetFirstChild<RunFonts>();
            if (existingFonts != null)
            {
                existingFonts.Ascii = primaryFont;
                existingFonts.HighAnsi = primaryFont;
                existingFonts.EastAsia = primaryFont;
                existingFonts.ComplexScript = primaryFont;
            }
            else
            {
                var newFonts = new RunFonts
                {
                    Ascii = primaryFont,
                    HighAnsi = primaryFont,
                    EastAsia = primaryFont,
                    ComplexScript = primaryFont
                };
                // Insert RunFonts at the correct position (after rStyle, near the beginning)
                InsertRunPropertyInOrder(runProps, newFonts);
            }
        }
    }

    /// <summary>
    /// Inserts a child element into RunProperties at the correct OOXML sequence position.
    /// OOXML order: rStyle, rFonts, b, bCs, i, iCs, caps, smallCaps, strike, dstrike, ...
    /// </summary>
    private static void InsertRunPropertyInOrder(RunProperties runProps, OpenXmlElement newElement)
    {
        // Define the correct order of rPr child elements (partial list of common elements)
        var elementOrder = new[]
        {
            typeof(RunStyle),           // rStyle
            typeof(RunFonts),           // rFonts
            typeof(Bold),               // b
            typeof(BoldComplexScript),  // bCs
            typeof(Italic),             // i
            typeof(ItalicComplexScript),// iCs
            typeof(Caps),               // caps
            typeof(SmallCaps),          // smallCaps
            typeof(Strike),             // strike
            typeof(DoubleStrike),       // dstrike
            typeof(Outline),            // outline
            typeof(Shadow),             // shadow
            typeof(Emboss),             // emboss
            typeof(Imprint),            // imprint
            typeof(NoProof),            // noProof
            typeof(SnapToGrid),         // snapToGrid
            typeof(Vanish),             // vanish
            typeof(WebHidden),          // webHidden
            typeof(Color),              // color
            typeof(Spacing),            // spacing
            typeof(CharacterScale),     // w
            typeof(Kern),               // kern
            typeof(Position),           // position
            typeof(FontSize),           // sz
            typeof(FontSizeComplexScript), // szCs
            typeof(Highlight),          // highlight
            typeof(Underline),          // u
            typeof(TextEffect),         // effect
            typeof(Border),             // bdr
            typeof(Shading),            // shd
            typeof(FitText),            // fitText
            typeof(VerticalTextAlignment), // vertAlign
            typeof(RightToLeftText),    // rtl
            typeof(ComplexScript),      // cs
            typeof(Emphasis),           // em
            typeof(Languages),          // lang
            typeof(EastAsianLayout),    // eastAsianLayout
            typeof(SpecVanish)          // specVanish
        };

        var newElementIndex = Array.IndexOf(elementOrder, newElement.GetType());
        if (newElementIndex < 0)
        {
            // Unknown element type, append at the end
            runProps.Append(newElement);
            return;
        }

        // Find the first element that should come after the new element
        foreach (var child in runProps.ChildElements)
        {
            var childIndex = Array.IndexOf(elementOrder, child.GetType());
            if (childIndex > newElementIndex)
            {
                // Insert before this child
                child.InsertBeforeSelf(newElement);
                return;
            }
        }

        // No element found that should come after, so append at the end
        runProps.Append(newElement);
    }

    /// <summary>
    /// Writes an image to the document with full formatting
    /// </summary>
    private void WriteImage(DocumentNode node)
    {
        var imageData = node.GetImageData();
        if (imageData?.Data == null || imageData.Data.Length == 0)
        {
            var placeholder = new Paragraph();
            placeholder.Append(new Run(new Text(node.Text.Length > 0 ? node.Text : "[Image]")));
            _body!.Append(placeholder);
            return;
        }

        string relationshipId;

        // Check if this image was restored from package data (use existing relationship)
        if (!string.IsNullOrEmpty(imageData.Id) && _imageRelationshipMapping.TryGetValue(imageData.Id, out var mappedRelId))
        {
            relationshipId = mappedRelId;
        }
        else
        {
            // Add new image part
            var imagePart = _mainPart!.AddImagePart(GetImageContentType(imageData.ContentType));
            using (var stream = new MemoryStream(imageData.Data))
            {
                imagePart.FeedData(stream);
            }
            relationshipId = _mainPart!.GetIdOfPart(imagePart);
        }

        // Use EMU dimensions for precise round-trip, or fall back to inches
        var widthEmu = imageData.WidthEmu > 0 ? imageData.WidthEmu : (long)(imageData.WidthInches * 914400);
        var heightEmu = imageData.HeightEmu > 0 ? imageData.HeightEmu : (long)(imageData.HeightInches * 914400);

        // Default size if not specified
        if (widthEmu <= 0) widthEmu = 914400 * 4;
        if (heightEmu <= 0) heightEmu = 914400 * 3;

        // Create the drawing element with formatting
        var drawing = CreateImageDrawing(relationshipId, widthEmu, heightEmu, imageData);

        var paragraph = new Paragraph();
        var run = new Run();
        run.Append(drawing);
        paragraph.Append(run);
        _body!.Append(paragraph);
    }

    /// <summary>
    /// Creates the Drawing element for an image with full formatting
    /// </summary>
    private Drawing CreateImageDrawing(string relationshipId, long widthEmu, long heightEmu, ImageData imageData)
    {
        var imageId = _imageCounter++;
        var fmt = imageData.Formatting;

        // Create common elements
        var extent = new DW.Extent { Cx = widthEmu, Cy = heightEmu };
        var docProperties = new DW.DocProperties
        {
            Id = imageId,
            Name = imageData.Name ?? $"Image{imageId}",
            Description = imageData.Description ?? ""
        };

        var graphic = new A.Graphic(
            new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = imageId, Name = imageData.Name ?? $"Image{imageId}" },
                            new PIC.NonVisualPictureDrawingProperties()
                        ),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId },
                            new A.Stretch(new A.FillRectangle())
                        ),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = widthEmu, Cy = heightEmu }
                            ),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                        )
                    )
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
        );

        // Create inline or anchor based on formatting
        if (fmt == null || fmt.IsInline)
        {
            var inline = new DW.Inline(
                extent,
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                docProperties,
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                graphic
            )
            {
                DistanceFromTop = (uint)(fmt?.DistanceFromTop ?? 0),
                DistanceFromBottom = (uint)(fmt?.DistanceFromBottom ?? 0),
                DistanceFromLeft = (uint)(fmt?.DistanceFromLeft ?? 0),
                DistanceFromRight = (uint)(fmt?.DistanceFromRight ?? 0)
            };

            return new Drawing(inline);
        }
        else
        {
            // Anchored image (more complex)
            var anchor = new DW.Anchor
            {
                DistanceFromTop = (uint)(fmt.DistanceFromTop ?? 0),
                DistanceFromBottom = (uint)(fmt.DistanceFromBottom ?? 0),
                DistanceFromLeft = (uint)(fmt.DistanceFromLeft ?? 0),
                DistanceFromRight = (uint)(fmt.DistanceFromRight ?? 0),
                SimplePos = false,
                RelativeHeight = (uint)(fmt.RelativeHeight ?? 0),
                BehindDoc = fmt.BehindDocument,
                Locked = fmt.Locked,
                LayoutInCell = fmt.LayoutInCell,
                AllowOverlap = fmt.AllowOverlap
            };

            anchor.Append(new DW.SimplePosition { X = 0, Y = 0 });

            // Horizontal position
            var hPosFrom = fmt.HorizontalRelativeTo switch
            {
                "Column" => DW.HorizontalRelativePositionValues.Column,
                "Page" => DW.HorizontalRelativePositionValues.Page,
                "Margin" => DW.HorizontalRelativePositionValues.Margin,
                _ => DW.HorizontalRelativePositionValues.Column
            };
            var hPos = new DW.HorizontalPosition { RelativeFrom = hPosFrom };
            if (long.TryParse(fmt.HorizontalPosition, out var hOffset))
            {
                hPos.Append(new DW.PositionOffset(hOffset.ToString()));
            }
            else
            {
                hPos.Append(new DW.PositionOffset("0"));
            }
            anchor.Append(hPos);

            // Vertical position
            var vPosFrom = fmt.VerticalRelativeTo switch
            {
                "Paragraph" => DW.VerticalRelativePositionValues.Paragraph,
                "Page" => DW.VerticalRelativePositionValues.Page,
                "Margin" => DW.VerticalRelativePositionValues.Margin,
                _ => DW.VerticalRelativePositionValues.Paragraph
            };
            var vPos = new DW.VerticalPosition { RelativeFrom = vPosFrom };
            if (long.TryParse(fmt.VerticalPosition, out var vOffset))
            {
                vPos.Append(new DW.PositionOffset(vOffset.ToString()));
            }
            else
            {
                vPos.Append(new DW.PositionOffset("0"));
            }
            anchor.Append(vPos);

            anchor.Append(extent);
            anchor.Append(new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 });

            // Wrap type
            switch (fmt.WrapType)
            {
                case "None":
                    anchor.Append(new DW.WrapNone());
                    break;
                case "Square":
                    anchor.Append(new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides });
                    break;
                case "Tight":
                    anchor.Append(new DW.WrapTight { WrapText = DW.WrapTextValues.BothSides });
                    break;
                case "TopAndBottom":
                    anchor.Append(new DW.WrapTopBottom());
                    break;
                default:
                    anchor.Append(new DW.WrapNone());
                    break;
            }

            anchor.Append(docProperties);
            anchor.Append(new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }));
            anchor.Append(graphic);

            return new Drawing(anchor);
        }
    }

    /// <summary>
    /// Gets the standardized image content type string
    /// </summary>
    private string GetImageContentType(string contentType)
    {
        return contentType.ToLowerInvariant() switch
        {
            "image/png" => "image/png",
            "image/gif" => "image/gif",
            "image/bmp" => "image/bmp",
            "image/tiff" => "image/tiff",
            "image/x-icon" or "image/vnd.microsoft.icon" => "image/x-icon",
            "image/x-emf" or "image/emf" => "image/x-emf",
            "image/x-wmf" or "image/wmf" => "image/x-wmf",
            _ => "image/jpeg"
        };
    }

    /// <summary>
    /// Writes a list container
    /// </summary>
    private void WriteList(DocumentNode node)
    {
        EnsureNumberingPart();

        foreach (var child in node.Children)
        {
            if (child.Type == ContentType.ListItem)
            {
                WriteListItem(child);
            }
            else
            {
                ProcessNode(child);
            }
        }

        _currentListId++;
    }

    /// <summary>
    /// Writes a list item
    /// </summary>
    private void WriteListItem(DocumentNode node)
    {
        EnsureNumberingPart();

        // Use original XML if available for exact round-trip
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            var updatedXml = UpdateImageRelationships(node.OriginalXml);
            var paragraph = new Paragraph(updatedXml);
            FixIndentationAttributes(paragraph);
            _body!.Append(paragraph);
        }
        else
        {
            var paragraph = new Paragraph();

            // Apply paragraph formatting
            var paragraphProps = CreateParagraphProperties(node);

            // Ensure list paragraph style and numbering
            if (paragraphProps.ParagraphStyleId == null)
            {
                paragraphProps.Append(new ParagraphStyleId { Val = "ListParagraph" });
            }

            var listLevel = 0;
            if (node.Metadata.TryGetValue("ListLevel", out var levelObj))
            {
                listLevel = Convert.ToInt32(levelObj);
            }

            if (paragraphProps.NumberingProperties == null)
            {
                paragraphProps.Append(new NumberingProperties(
                    new NumberingLevelReference { Val = listLevel },
                    new NumberingId { Val = _currentListId }
                ));
            }

            paragraph.Append(paragraphProps);

            // Write formatted runs or plain text
            WriteRunsOrText(paragraph, node);

            _body!.Append(paragraph);
        }

        // Process nested children
        foreach (var child in node.Children)
        {
            ProcessNode(child);
        }
    }

    /// <summary>
    /// Ensures the numbering part exists for lists
    /// </summary>
    private void EnsureNumberingPart()
    {
        if (_numberingPart != null) return;

        _numberingPart = _mainPart!.AddNewPart<NumberingDefinitionsPart>();

        var numbering = new Numbering();

        var abstractNum = new AbstractNum { AbstractNumberId = 0 };
        abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        var bullets = new[] { "●", "○", "■", "□", "◆", "◇", "▸", "▹", "•" };
        for (int i = 0; i < 9; i++)
        {
            var level = new Level { LevelIndex = i };
            level.Append(new StartNumberingValue { Val = 1 });
            level.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
            level.Append(new LevelText { Val = bullets[i] });
            level.Append(new LevelJustification { Val = LevelJustificationValues.Left });
            level.Append(new PreviousParagraphProperties(
                new Indentation { Left = ((i + 1) * 720).ToString(), Hanging = "360" }
            ));
            abstractNum.Append(level);
        }

        numbering.Append(abstractNum);

        for (int i = 1; i <= 10; i++)
        {
            numbering.Append(new NumberingInstance(
                    new AbstractNumId { Val = 0 }
                )
                { NumberID = i });
        }

        _numberingPart.Numbering = numbering;
    }

    /// <summary>
    /// Adds section properties to the document
    /// </summary>
    private void AddSectionProperties()
    {
        // Restore original section properties if available
        if (_packageData?.SectionPropertiesXml?.Count > 0)
        {
            // Use the last section properties (main document section)
            var lastSectPrXml = _packageData.SectionPropertiesXml.Last();
            // Clean the section properties XML to remove problematic attributes
            var cleanedSectPrXml = CleanXmlAttributes(lastSectPrXml);
            var sectionProps = new SectionProperties(cleanedSectPrXml);

            // Update header/footer references if they were restored
            UpdateSectionHeaderFooterReferences(sectionProps);

            _body!.Append(sectionProps);
        }
        else
        {
            var sectionProps = new SectionProperties();
            sectionProps.Append(new PageSize { Width = 12240, Height = 15840 });
            sectionProps.Append(new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 });
            _body!.Append(sectionProps);
        }
    }

    /// <summary>
    /// Updates header and footer references in section properties to use new relationship IDs
    /// </summary>
    private void UpdateSectionHeaderFooterReferences(SectionProperties sectionProps)
    {
        // Update header references using the relationship mapping
        foreach (var headerRef in sectionProps.Elements<HeaderReference>())
        {
            var oldId = headerRef.Id?.Value;
            if (!string.IsNullOrEmpty(oldId) && _headerRelationshipMapping.TryGetValue(oldId, out var newId))
            {
                headerRef.Id = newId;
            }
        }

        // Update footer references using the relationship mapping
        foreach (var footerRef in sectionProps.Elements<FooterReference>())
        {
            var oldId = footerRef.Id?.Value;
            if (!string.IsNullOrEmpty(oldId) && _footerRelationshipMapping.TryGetValue(oldId, out var newId))
            {
                footerRef.Id = newId;
            }
        }
    }

    /// <summary>Releases resources used by the writer.</summary>
    public void Dispose()
    {
        _document?.Dispose();
        _document = null;
        _mainPart = null;
        _body = null;
    }
}
