using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WpTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WpTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace WordDocumentParser;

/// <summary>
/// Parses Word documents and builds a hierarchical tree structure based on headings.
/// Captures full formatting and document package data for round-trip fidelity.
/// </summary>
public class WordDocumentTreeParser : IDisposable
{
    private WordprocessingDocument? _document;
    private MainDocumentPart? _mainPart;
    private readonly Dictionary<string, Style> _styleCache = [];
    private readonly Dictionary<string, string> _hyperlinkUrls = [];

    /// <summary>
    /// Parses a Word document from a file path and returns the document tree
    /// </summary>
    public DocumentNode ParseFromFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Document not found: {filePath}");

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return ParseFromStream(stream, Path.GetFileName(filePath));
    }

    /// <summary>
    /// Parses a Word document from a stream and returns the document tree
    /// </summary>
    public DocumentNode ParseFromStream(Stream stream, string documentName = "Document")
    {
        _document = WordprocessingDocument.Open(stream, false);
        _mainPart = _document.MainDocumentPart;

        if (_mainPart?.Document?.Body is null)
            throw new InvalidOperationException("Invalid Word document: missing body");

        // Cache styles and hyperlink relationships
        CacheStyles();
        CacheHyperlinkRelationships();

        // Create root node
        var root = new DocumentNode(ContentType.Document, documentName)
        {
            Metadata = { ["FileName"] = documentName }
        };

        // Extract and store all document package data for round-trip fidelity
        root.PackageData = ExtractPackageData();

        // Parse all body elements
        var bodyElements = _mainPart.Document.Body.Elements().ToList();
        BuildTree(root, bodyElements);

        return root;
    }

    /// <summary>
    /// Extracts all document package data for round-trip preservation
    /// </summary>
    private DocumentPackageData ExtractPackageData()
    {
        var packageData = new DocumentPackageData();

        // Store original document XML
        if (_mainPart?.Document is not null)
        {
            packageData.OriginalDocumentXml = _mainPart.Document.OuterXml;
        }

        // Extract styles
        if (_mainPart?.StyleDefinitionsPart?.Styles is not null)
        {
            packageData.StylesXml = _mainPart.StyleDefinitionsPart.Styles.OuterXml;
        }

        // Extract theme
        if (_mainPart?.ThemePart?.Theme is not null)
        {
            packageData.ThemeXml = _mainPart.ThemePart.Theme.OuterXml;
        }

        // Extract font table
        if (_mainPart?.FontTablePart?.Fonts is not null)
        {
            packageData.FontTableXml = _mainPart.FontTablePart.Fonts.OuterXml;
        }

        // Extract numbering definitions
        if (_mainPart?.NumberingDefinitionsPart?.Numbering is not null)
        {
            packageData.NumberingXml = _mainPart.NumberingDefinitionsPart.Numbering.OuterXml;
        }

        // Extract document settings
        if (_mainPart?.DocumentSettingsPart?.Settings is not null)
        {
            packageData.SettingsXml = _mainPart.DocumentSettingsPart.Settings.OuterXml;
        }

        // Extract web settings
        if (_mainPart?.WebSettingsPart?.WebSettings is not null)
        {
            packageData.WebSettingsXml = _mainPart.WebSettingsPart.WebSettings.OuterXml;
        }

        // Extract footnotes
        if (_mainPart?.FootnotesPart?.Footnotes is not null)
        {
            packageData.FootnotesXml = _mainPart.FootnotesPart.Footnotes.OuterXml;
        }

        // Extract endnotes
        if (_mainPart?.EndnotesPart?.Endnotes is not null)
        {
            packageData.EndnotesXml = _mainPart.EndnotesPart.Endnotes.OuterXml;
        }

        // Extract headers
        foreach (var headerPart in _mainPart?.HeaderParts ?? [])
        {
            var relId = _mainPart!.GetIdOfPart(headerPart);
            if (headerPart.Header is not null)
            {
                packageData.Headers[relId] = headerPart.Header.OuterXml;
            }
        }

        // Extract footers
        foreach (var footerPart in _mainPart?.FooterParts ?? [])
        {
            var relId = _mainPart!.GetIdOfPart(footerPart);
            if (footerPart.Footer is not null)
            {
                packageData.Footers[relId] = footerPart.Footer.OuterXml;
            }
        }

        // Extract images
        foreach (var imagePart in _mainPart?.ImageParts ?? [])
        {
            var relId = _mainPart!.GetIdOfPart(imagePart);
            try
            {
                using var imgStream = imagePart.GetStream();
                using var ms = new MemoryStream();
                imgStream.CopyTo(ms);
                packageData.Images[relId] = new ImagePartData
                {
                    ContentType = imagePart.ContentType,
                    Data = ms.ToArray(),
                    OriginalRelationshipId = relId,
                    OriginalUri = imagePart.Uri?.ToString()
                };
            }
            catch
            {
                // Skip images that can't be read
            }
        }

        // Extract hyperlink relationships
        foreach (var rel in _mainPart?.HyperlinkRelationships ?? [])
        {
            packageData.HyperlinkRelationships[rel.Id] = new HyperlinkRelationshipData
            {
                Url = rel.Uri.ToString(),
                IsExternal = rel.IsExternal
            };
        }

        // Extract core properties
        if (_document?.PackageProperties is not null)
        {
            var props = _document.PackageProperties;
            packageData.CoreProperties = new CoreProperties
            {
                Title = props.Title,
                Subject = props.Subject,
                Creator = props.Creator,
                Keywords = props.Keywords,
                Description = props.Description,
                LastModifiedBy = props.LastModifiedBy,
                Revision = props.Revision,
                Created = props.Created?.ToString("o"),
                Modified = props.Modified?.ToString("o"),
                Category = props.Category,
                ContentStatus = props.ContentStatus
            };
        }

        // Extract extended properties
        if (_document?.ExtendedFilePropertiesPart?.Properties is not null)
        {
            var extProps = _document.ExtendedFilePropertiesPart.Properties;
            packageData.ExtendedProperties = new ExtendedProperties
            {
                Template = extProps.Template?.Text,
                Application = extProps.Application?.Text,
                AppVersion = extProps.ApplicationVersion?.Text,
                Company = extProps.Company?.Text,
                Manager = extProps.Manager?.Text
            };

            if (int.TryParse(extProps.Pages?.Text, out var pages))
                packageData.ExtendedProperties.Pages = pages;
            if (int.TryParse(extProps.Words?.Text, out var words))
                packageData.ExtendedProperties.Words = words;
            if (int.TryParse(extProps.Characters?.Text, out var chars))
                packageData.ExtendedProperties.Characters = chars;
            if (int.TryParse(extProps.CharactersWithSpaces?.Text, out var charsSpaces))
                packageData.ExtendedProperties.CharactersWithSpaces = charsSpaces;
            if (int.TryParse(extProps.Lines?.Text, out var lines))
                packageData.ExtendedProperties.Lines = lines;
            if (int.TryParse(extProps.Paragraphs?.Text, out var paras))
                packageData.ExtendedProperties.Paragraphs = paras;
            if (int.TryParse(extProps.TotalTime?.Text, out var time))
                packageData.ExtendedProperties.TotalTime = time;
        }

        // Extract custom properties XML
        if (_document?.CustomFilePropertiesPart?.Properties is not null)
        {
            packageData.CustomPropertiesXml = _document.CustomFilePropertiesPart.Properties.OuterXml;
        }

        // Extract section properties from body
        var body = _mainPart?.Document?.Body;
        if (body is not null)
        {
            foreach (var sectPr in body.Descendants<SectionProperties>())
            {
                packageData.SectionPropertiesXml.Add(sectPr.OuterXml);
            }
        }

        // Extract raw core properties XML for exact round-trip
        if (_document?.CoreFilePropertiesPart is not null)
        {
            using var stream = _document.CoreFilePropertiesPart.GetStream();
            using var reader = new StreamReader(stream);
            packageData.CorePropertiesXml = reader.ReadToEnd();
        }

        // Extract raw app properties XML for exact round-trip
        if (_document?.ExtendedFilePropertiesPart is not null)
        {
            using var stream = _document.ExtendedFilePropertiesPart.GetStream();
            using var reader = new StreamReader(stream);
            packageData.AppPropertiesXml = reader.ReadToEnd();
        }

        // Extract custom XML parts
        var customXmlParts = _mainPart?.Parts
            .Where(p => p.OpenXmlPart is CustomXmlPart)
            .Select(p => p.OpenXmlPart as CustomXmlPart) ?? [];

        foreach (var customXmlPart in customXmlParts)
        {
            if (customXmlPart is null) continue;
            try
            {
                using var stream = customXmlPart.GetStream();
                using var reader = new StreamReader(stream);
                var xmlContent = reader.ReadToEnd();

                string? propsXml = null;
                if (customXmlPart.CustomXmlPropertiesPart is not null)
                {
                    using var propsStream = customXmlPart.CustomXmlPropertiesPart.GetStream();
                    using var propsReader = new StreamReader(propsStream);
                    propsXml = propsReader.ReadToEnd();
                }

                packageData.CustomXmlParts[customXmlPart.Uri.ToString()] = new CustomXmlPartData
                {
                    XmlContent = xmlContent,
                    PropertiesXml = propsXml
                };
            }
            catch
            {
                // Skip custom XML parts that can't be read
            }
        }

        // Extract Glossary Document Part (for Quick Parts, building blocks, document property fields)
        if (_mainPart?.GlossaryDocumentPart is not null)
        {
            var glossaryPart = _mainPart.GlossaryDocumentPart;
            if (glossaryPart.GlossaryDocument is not null)
            {
                packageData.GlossaryDocumentXml = glossaryPart.GlossaryDocument.OuterXml;
            }

            // Extract glossary styles
            if (glossaryPart.StyleDefinitionsPart?.Styles is not null)
            {
                packageData.GlossaryStylesXml = glossaryPart.StyleDefinitionsPart.Styles.OuterXml;
            }

            // Extract glossary font table
            if (glossaryPart.FontTablePart?.Fonts is not null)
            {
                packageData.GlossaryFontTableXml = glossaryPart.FontTablePart.Fonts.OuterXml;
            }

            // Extract glossary images
            foreach (var imagePart in glossaryPart.ImageParts ?? [])
            {
                var relId = glossaryPart.GetIdOfPart(imagePart);
                try
                {
                    using var imgStream = imagePart.GetStream();
                    using var ms = new MemoryStream();
                    imgStream.CopyTo(ms);
                    packageData.GlossaryImages[relId] = new ImagePartData
                    {
                        ContentType = imagePart.ContentType,
                        Data = ms.ToArray(),
                        OriginalRelationshipId = relId,
                        OriginalUri = imagePart.Uri?.ToString()
                    };
                }
                catch
                {
                    // Skip images that can't be read
                }
            }
        }

        return packageData;
    }

    /// <summary>
    /// Caches document styles for efficient heading level detection
    /// </summary>
    private void CacheStyles()
    {
        _styleCache.Clear();
        var stylesPart = _mainPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles is null) return;

        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            var styleId = style.StyleId?.Value;
            if (!string.IsNullOrEmpty(styleId))
            {
                _styleCache[styleId] = style;
            }
        }
    }

    /// <summary>
    /// Caches hyperlink relationships to resolve URLs
    /// </summary>
    private void CacheHyperlinkRelationships()
    {
        _hyperlinkUrls.Clear();
        if (_mainPart is null) return;

        foreach (var rel in _mainPart.HyperlinkRelationships)
        {
            _hyperlinkUrls[rel.Id] = rel.Uri.ToString();
        }
    }

    /// <summary>
    /// Builds the tree structure by processing body elements and organizing by heading hierarchy
    /// </summary>
    private void BuildTree(DocumentNode root, List<OpenXmlElement> elements)
    {
        const int maxLevel = 9;
        var stack = new DocumentNode?[maxLevel + 1];
        stack[0] = root;
        var currentLevel = 0;

        foreach (var element in elements)
        {
            var node = ProcessElement(element);
            if (node is null) continue;

            if (node.Type == ContentType.Heading && node.HeadingLevel > 0)
            {
                var level = Math.Min(node.HeadingLevel, maxLevel);
                var parentLevel = Math.Min(currentLevel, level - 1);
                while (parentLevel > 0 && stack[parentLevel] is null)
                    parentLevel--;

                var parent = stack[parentLevel] ?? root;
                parent.AddChild(node);
                stack[level] = node;
                currentLevel = level;
            }
            else
            {
                var container = stack[currentLevel] ?? root;
                container.AddChild(node);
            }
        }
    }

    /// <summary>
    /// Processes a single OpenXML element and returns a DocumentNode
    /// </summary>
    private DocumentNode? ProcessElement(OpenXmlElement element) => element switch
    {
        Paragraph para => ProcessParagraph(para),
        Table table => ProcessTable(table),
        SdtBlock sdtBlock => ProcessStructuredDocumentTag(sdtBlock),
        _ => null
    };

    /// <summary>
    /// Processes a paragraph element with full formatting capture
    /// </summary>
    private DocumentNode? ProcessParagraph(Paragraph para)
    {
        var runs = ExtractFormattedRuns(para);
        var text = string.Concat(runs.Select(r => r.IsTab ? "\t" : r.IsBreak ? " " : r.Text)).Trim();
        var headingLevel = GetHeadingLevel(para);

        // Check for complex content that should be preserved exactly
        var hasComplexContent = para.Descendants<AlternateContent>().Any() ||
                                para.OuterXml.Contains("wpc:") ||
                                para.OuterXml.Contains("v:group") ||
                                para.OuterXml.Contains("v:shape");

        // Check for section properties (section breaks) that must be preserved
        var hasSectionProperties = para.ParagraphProperties?.SectionProperties is not null;

        // Check for field characters (TOC, cross-references, page numbers, etc.)
        var hasFieldCharacters = para.Descendants<FieldChar>().Any() ||
                                 para.Descendants<FieldCode>().Any();

        // Skip empty paragraphs (but keep empty headings, complex content, section breaks, and field characters)
        if (string.IsNullOrWhiteSpace(text) && headingLevel == 0 && runs.Count == 0 &&
            !hasComplexContent && !hasSectionProperties && !hasFieldCharacters)
            return null;

        DocumentNode node;

        if (headingLevel > 0)
        {
            node = new DocumentNode(ContentType.Heading, headingLevel, text);
        }
        else if (IsListParagraph(para))
        {
            node = new DocumentNode(ContentType.ListItem, text);
            var numPr = para.ParagraphProperties?.NumberingProperties;
            if (numPr is not null)
            {
                node.Metadata["ListLevel"] = numPr.NumberingLevelReference?.Val?.Value ?? 0;
                node.Metadata["ListId"] = numPr.NumberingId?.Val?.Value ?? 0;
            }
        }
        else
        {
            node = new DocumentNode(ContentType.Paragraph, text);
        }

        // Store the original XML for exact round-trip fidelity
        node.OriginalXml = para.OuterXml;

        // Store formatted runs
        node.Runs = runs;

        // Capture paragraph formatting
        node.ParagraphFormatting = ExtractParagraphFormatting(para);

        // Check for images in the paragraph
        var images = ExtractImages(para);
        foreach (var imageNode in images)
        {
            node.AddChild(imageNode);
        }

        // Extract hyperlinks with URLs
        var hyperlinks = ExtractHyperlinks(para);
        if (hyperlinks.Count > 0)
        {
            node.Metadata["HasHyperlinks"] = true;
            node.Metadata["Hyperlinks"] = hyperlinks;
        }

        return node;
    }

    /// <summary>
    /// Extracts formatted runs from a paragraph, preserving all text formatting
    /// </summary>
    private List<FormattedRun> ExtractFormattedRuns(Paragraph para)
    {
        List<FormattedRun> formattedRuns = [];

        foreach (var child in para.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    formattedRuns.AddRange(ProcessRun(run));
                    break;
                case Hyperlink hyperlink:
                    formattedRuns.AddRange(ProcessHyperlinkRuns(hyperlink));
                    break;
                case BookmarkStart:
                case BookmarkEnd:
                case ProofError:
                    // Skip these elements
                    break;
            }
        }

        return formattedRuns;
    }

    /// <summary>
    /// Processes a single run element and extracts formatted runs
    /// </summary>
    private List<FormattedRun> ProcessRun(Run run)
    {
        List<FormattedRun> result = [];
        var formatting = ExtractRunFormatting(run.RunProperties);

        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case Text text:
                    result.Add(new FormattedRun(text.Text, formatting.Clone()));
                    break;
                case TabChar:
                    result.Add(new FormattedRun { IsTab = true, Formatting = formatting.Clone() });
                    break;
                case Break br:
                    result.Add(new FormattedRun
                    {
                        IsBreak = true,
                        BreakType = br.Type?.Value.ToString(),
                        Formatting = formatting.Clone()
                    });
                    break;
                case CarriageReturn:
                    result.Add(new FormattedRun { IsBreak = true, BreakType = "CarriageReturn", Formatting = formatting.Clone() });
                    break;
            }
        }

        return result;
    }

    /// <summary>
    /// Processes runs within a hyperlink
    /// </summary>
    private List<FormattedRun> ProcessHyperlinkRuns(Hyperlink hyperlink)
    {
        List<FormattedRun> result = [];

        foreach (var run in hyperlink.Elements<Run>())
        {
            var runs = ProcessRun(run);
            // Mark these runs as being part of a hyperlink
            foreach (var r in runs)
            {
                r.Formatting.StyleId = "Hyperlink";
            }
            result.AddRange(runs);
        }

        return result;
    }

    /// <summary>
    /// Extracts run formatting properties
    /// </summary>
    private static RunFormatting ExtractRunFormatting(RunProperties? rPr)
    {
        var formatting = new RunFormatting();
        if (rPr is null) return formatting;

        // Bold
        formatting.Bold = rPr.Bold is not null && (rPr.Bold.Val is null || rPr.Bold.Val.Value);

        // Italic
        formatting.Italic = rPr.Italic is not null && (rPr.Italic.Val is null || rPr.Italic.Val.Value);

        // Underline
        if (rPr.Underline is not null)
        {
            formatting.Underline = rPr.Underline.Val?.Value != UnderlineValues.None;
            formatting.UnderlineStyle = rPr.Underline.Val?.Value.ToString();
        }

        // Strike
        formatting.Strike = rPr.Strike is not null && (rPr.Strike.Val is null || rPr.Strike.Val.Value);
        formatting.DoubleStrike = rPr.DoubleStrike is not null && (rPr.DoubleStrike.Val is null || rPr.DoubleStrike.Val.Value);

        // Font
        var fonts = rPr.RunFonts;
        if (fonts is not null)
        {
            formatting.FontFamily = fonts.HighAnsi?.Value;
            formatting.FontFamilyAscii = fonts.Ascii?.Value;
            formatting.FontFamilyEastAsia = fonts.EastAsia?.Value;
            formatting.FontFamilyComplexScript = fonts.ComplexScript?.Value;
        }

        // Font size
        formatting.FontSize = rPr.FontSize?.Val?.Value;
        formatting.FontSizeComplexScript = rPr.FontSizeComplexScript?.Val?.Value;

        // Color
        formatting.Color = rPr.Color?.Val?.Value;

        // Highlight
        formatting.Highlight = rPr.Highlight?.Val?.Value.ToString();

        // Superscript/Subscript
        var vertAlign = rPr.VerticalTextAlignment?.Val?.Value;
        if (vertAlign.HasValue)
        {
            formatting.Superscript = vertAlign.Value == VerticalPositionValues.Superscript;
            formatting.Subscript = vertAlign.Value == VerticalPositionValues.Subscript;
        }

        // Caps
        formatting.SmallCaps = rPr.SmallCaps is not null && (rPr.SmallCaps.Val is null || rPr.SmallCaps.Val.Value);
        formatting.AllCaps = rPr.Caps is not null && (rPr.Caps.Val is null || rPr.Caps.Val.Value);

        // Shading
        formatting.Shading = rPr.Shading?.Fill?.Value;

        // Character style
        formatting.StyleId = rPr.RunStyle?.Val?.Value;

        return formatting;
    }

    /// <summary>
    /// Extracts paragraph formatting properties
    /// </summary>
    private ParagraphFormatting ExtractParagraphFormatting(Paragraph para)
    {
        var formatting = new ParagraphFormatting();
        var pPr = para.ParagraphProperties;
        if (pPr is null) return formatting;

        // Style
        formatting.StyleId = pPr.ParagraphStyleId?.Val?.Value;

        // Alignment
        formatting.Alignment = pPr.Justification?.Val?.Value.ToString();

        // Indentation
        var ind = pPr.Indentation;
        if (ind is not null)
        {
            formatting.IndentLeft = ind.Left?.Value;
            formatting.IndentRight = ind.Right?.Value;
            formatting.IndentFirstLine = ind.FirstLine?.Value;
            formatting.IndentHanging = ind.Hanging?.Value;
        }

        // Spacing
        var spacing = pPr.SpacingBetweenLines;
        if (spacing is not null)
        {
            formatting.SpacingBefore = spacing.Before?.Value;
            formatting.SpacingAfter = spacing.After?.Value;
            formatting.LineSpacing = spacing.Line?.Value;
            formatting.LineSpacingRule = spacing.LineRule?.Value.ToString();
        }

        // Keep with next/keep lines
        formatting.KeepNext = pPr.KeepNext is not null;
        formatting.KeepLines = pPr.KeepLines is not null;

        // Page break before
        formatting.PageBreakBefore = pPr.PageBreakBefore is not null;

        // Widow control
        formatting.WidowControl = pPr.WidowControl is not null;

        // Outline level
        formatting.OutlineLevel = pPr.OutlineLevel?.Val?.Value.ToString();

        // Shading
        var shading = pPr.Shading;
        if (shading is not null)
        {
            formatting.ShadingFill = shading.Fill?.Value;
            formatting.ShadingColor = shading.Color?.Value;
        }

        // Borders
        var borders = pPr.ParagraphBorders;
        if (borders is not null)
        {
            formatting.TopBorder = ExtractBorderFormatting(borders.TopBorder);
            formatting.BottomBorder = ExtractBorderFormatting(borders.BottomBorder);
            formatting.LeftBorder = ExtractBorderFormatting(borders.LeftBorder);
            formatting.RightBorder = ExtractBorderFormatting(borders.RightBorder);
        }

        // Numbering
        var numPr = pPr.NumberingProperties;
        if (numPr is not null)
        {
            formatting.NumberingId = numPr.NumberingId?.Val?.Value;
            formatting.NumberingLevel = numPr.NumberingLevelReference?.Val?.Value;
        }

        return formatting;
    }

    /// <summary>
    /// Extracts border formatting
    /// </summary>
    private static BorderFormatting? ExtractBorderFormatting(BorderType? border)
    {
        if (border is null) return null;

        return new BorderFormatting
        {
            Style = border.Val?.Value.ToString(),
            Size = border.Size?.Value.ToString(),
            Color = border.Color?.Value,
            Space = border.Space?.Value.ToString()
        };
    }

    /// <summary>
    /// Extracts hyperlinks with their URLs
    /// </summary>
    private List<HyperlinkData> ExtractHyperlinks(Paragraph para)
    {
        List<HyperlinkData> result = [];

        foreach (var hyperlink in para.Descendants<Hyperlink>())
        {
            var data = new HyperlinkData
            {
                RelationshipId = hyperlink.Id?.Value,
                Anchor = hyperlink.Anchor?.Value,
                Tooltip = hyperlink.Tooltip?.Value
            };

            // Get URL from relationship
            if (!string.IsNullOrEmpty(data.RelationshipId) && _hyperlinkUrls.TryGetValue(data.RelationshipId, out var url))
            {
                data.Url = url;
            }

            // Extract text and runs
            List<FormattedRun> runs = [];
            foreach (var run in hyperlink.Elements<Run>())
            {
                runs.AddRange(ProcessRun(run));
            }
            data.Runs = runs;
            data.Text = string.Concat(runs.Select(r => r.Text));

            result.Add(data);
        }

        return result;
    }

    /// <summary>
    /// Extracts the heading level from a paragraph (0 if not a heading)
    /// </summary>
    private int GetHeadingLevel(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(styleId)) return 0;

        if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
        {
            var levelStr = styleId[7..];
            if (int.TryParse(levelStr, out var level) && level is >= 1 and <= 9)
                return level;
        }

        if (_styleCache.TryGetValue(styleId, out var style))
        {
            var outlineLevel = style.StyleParagraphProperties?.OutlineLevel?.Val;
            if (outlineLevel is not null)
                return outlineLevel.Value + 1;

            var basedOn = style.BasedOn?.Val?.Value;
            if (!string.IsNullOrEmpty(basedOn) && basedOn.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            {
                var levelStr = basedOn[7..];
                if (int.TryParse(levelStr, out var level) && level is >= 1 and <= 9)
                    return level;
            }
        }

        var directOutline = para.ParagraphProperties?.OutlineLevel?.Val;
        if (directOutline is not null)
            return directOutline.Value + 1;

        return 0;
    }

    /// <summary>
    /// Checks if a paragraph is part of a list
    /// </summary>
    private static bool IsListParagraph(Paragraph para) => para.ParagraphProperties?.NumberingProperties is not null;

    /// <summary>
    /// Processes a table element with full formatting capture
    /// </summary>
    private DocumentNode ProcessTable(Table table)
    {
        var node = new DocumentNode(ContentType.Table, "[Table]");
        var tableData = new TableData();

        // Extract table formatting
        tableData.Formatting = ExtractTableFormatting(table);

        // Extract grid column widths
        var grid = table.GetFirstChild<TableGrid>();
        if (grid is not null)
        {
            tableData.Formatting ??= new TableFormatting();
            tableData.Formatting.GridColumnWidths = [.. grid.Elements<GridColumn>().Select(c => c.Width?.Value ?? "")];
        }

        var rowIndex = 0;
        foreach (var row in table.Elements<WpTableRow>())
        {
            var tableRow = new TableRow { RowIndex = rowIndex };

            // Extract row formatting
            tableRow.Formatting = ExtractTableRowFormatting(row);
            tableRow.IsHeader = tableRow.Formatting?.IsHeader ?? false;

            var colIndex = 0;
            foreach (var cell in row.Elements<WpTableCell>())
            {
                var tableCell = new TableCell
                {
                    RowIndex = rowIndex,
                    ColumnIndex = colIndex
                };

                // Extract cell formatting
                tableCell.Formatting = ExtractTableCellFormatting(cell);

                // Apply span values from formatting
                if (tableCell.Formatting is not null)
                {
                    tableCell.ColSpan = tableCell.Formatting.GridSpan;
                    if (tableCell.Formatting.VerticalMerge == "Restart")
                        tableCell.RowSpan = -1;
                    else if (tableCell.Formatting.VerticalMerge == "Continue")
                        tableCell.RowSpan = 0;
                }

                // Process cell content
                foreach (var para in cell.Elements<Paragraph>())
                {
                    var paraNode = ProcessParagraph(para);
                    if (paraNode is not null)
                    {
                        tableCell.Content.Add(paraNode);
                    }
                }

                tableRow.Cells.Add(tableCell);
                colIndex += tableCell.ColSpan;
            }

            if (colIndex > tableData.ColumnCount)
                tableData.ColumnCount = colIndex;

            tableData.Rows.Add(tableRow);
            rowIndex++;
        }

        node.Metadata["TableData"] = tableData;
        node.Metadata["RowCount"] = tableData.RowCount;
        node.Metadata["ColumnCount"] = tableData.ColumnCount;
        node.Text = $"[Table: {tableData.RowCount}x{tableData.ColumnCount}]";

        // Store the original XML for exact round-trip fidelity
        node.OriginalXml = table.OuterXml;

        return node;
    }

    /// <summary>
    /// Extracts table formatting properties
    /// </summary>
    private static TableFormatting ExtractTableFormatting(Table table)
    {
        var formatting = new TableFormatting();
        var tblPr = table.GetFirstChild<TableProperties>();
        if (tblPr is null) return formatting;

        // Width
        var width = tblPr.TableWidth;
        if (width is not null)
        {
            formatting.Width = width.Width?.Value;
            formatting.WidthType = width.Type?.Value.ToString();
        }

        // Alignment
        formatting.Alignment = tblPr.TableJustification?.Val?.Value.ToString();

        // Indent
        formatting.IndentFromLeft = tblPr.TableIndentation?.Width?.Value.ToString();

        // Borders
        var borders = tblPr.TableBorders;
        if (borders is not null)
        {
            formatting.TopBorder = ExtractBorderFormatting(borders.TopBorder);
            formatting.BottomBorder = ExtractBorderFormatting(borders.BottomBorder);
            formatting.LeftBorder = ExtractBorderFormatting(borders.LeftBorder);
            formatting.RightBorder = ExtractBorderFormatting(borders.RightBorder);
            formatting.InsideHorizontalBorder = ExtractBorderFormatting(borders.InsideHorizontalBorder);
            formatting.InsideVerticalBorder = ExtractBorderFormatting(borders.InsideVerticalBorder);
        }

        // Cell margins
        var margins = tblPr.TableCellMarginDefault;
        if (margins is not null)
        {
            formatting.CellMarginTop = margins.TopMargin?.Width?.Value;
            formatting.CellMarginBottom = margins.BottomMargin?.Width?.Value;
            formatting.CellMarginLeft = margins.TableCellLeftMargin?.Width?.Value.ToString();
            formatting.CellMarginRight = margins.TableCellRightMargin?.Width?.Value.ToString();
        }

        return formatting;
    }

    /// <summary>
    /// Extracts table row formatting properties
    /// </summary>
    private static TableRowFormatting ExtractTableRowFormatting(WpTableRow row)
    {
        var formatting = new TableRowFormatting();
        var trPr = row.TableRowProperties;
        if (trPr is null) return formatting;

        // Height
        var height = trPr.GetFirstChild<TableRowHeight>();
        if (height is not null)
        {
            formatting.Height = height.Val?.Value.ToString();
            formatting.HeightRule = height.HeightType?.Value.ToString();
        }

        // Header
        formatting.IsHeader = trPr.GetFirstChild<TableHeader>() is not null;

        // Can't split
        formatting.CantSplit = trPr.GetFirstChild<CantSplit>() is not null;

        return formatting;
    }

    /// <summary>
    /// Extracts table cell formatting properties
    /// </summary>
    private static TableCellFormatting ExtractTableCellFormatting(WpTableCell cell)
    {
        var formatting = new TableCellFormatting();
        var tcPr = cell.TableCellProperties;
        if (tcPr is null) return formatting;

        // Width
        var width = tcPr.TableCellWidth;
        if (width is not null)
        {
            formatting.Width = width.Width?.Value;
            formatting.WidthType = width.Type?.Value.ToString();
        }

        // Grid span
        formatting.GridSpan = (int)(tcPr.GridSpan?.Val?.Value ?? 1);

        // Vertical merge
        var vMerge = tcPr.VerticalMerge;
        if (vMerge is not null)
        {
            formatting.VerticalMerge = vMerge.Val?.Value == MergedCellValues.Restart ? "Restart" : "Continue";
        }

        // Vertical alignment
        formatting.VerticalAlignment = tcPr.TableCellVerticalAlignment?.Val?.Value.ToString();

        // Shading
        var shading = tcPr.Shading;
        if (shading is not null)
        {
            formatting.ShadingFill = shading.Fill?.Value;
            formatting.ShadingColor = shading.Color?.Value;
            formatting.ShadingPattern = shading.Val?.Value.ToString();
        }

        // Borders
        var borders = tcPr.TableCellBorders;
        if (borders is not null)
        {
            formatting.TopBorder = ExtractBorderFormatting(borders.TopBorder);
            formatting.BottomBorder = ExtractBorderFormatting(borders.BottomBorder);
            formatting.LeftBorder = ExtractBorderFormatting(borders.LeftBorder);
            formatting.RightBorder = ExtractBorderFormatting(borders.RightBorder);
        }

        // Text direction
        formatting.TextDirection = tcPr.TextDirection?.Val?.Value.ToString();

        // No wrap
        formatting.NoWrap = tcPr.NoWrap is not null;

        return formatting;
    }

    /// <summary>
    /// Extracts images from a paragraph
    /// </summary>
    private List<DocumentNode> ExtractImages(Paragraph para)
    {
        List<DocumentNode> images = [];

        var drawings = para.Descendants<Drawing>().ToList();
        foreach (var drawing in drawings)
        {
            var imageNode = ProcessDrawing(drawing);
            if (imageNode is not null)
                images.Add(imageNode);
        }

        return images;
    }

    /// <summary>
    /// Processes a drawing element to extract image information with full formatting
    /// </summary>
    private DocumentNode? ProcessDrawing(Drawing drawing)
    {
        var inline = drawing.Inline;
        var anchor = drawing.Anchor;

        var extent = inline?.Extent ?? anchor?.GetFirstChild<DW.Extent>();
        var docPr = inline?.DocProperties ?? anchor?.GetFirstChild<DW.DocProperties>();
        var graphic = inline?.Graphic ?? anchor?.GetFirstChild<A.Graphic>();

        if (graphic is null) return null;

        var blip = graphic.Descendants<A.Blip>().FirstOrDefault();
        if (blip is null) return null;

        var imageData = new ImageData();

        // Get image relationship ID and data
        var embedId = blip.Embed?.Value;
        if (!string.IsNullOrEmpty(embedId) && _mainPart is not null)
        {
            imageData.Id = embedId;

            try
            {
                if (_mainPart.GetPartById(embedId) is ImagePart imagePart)
                {
                    imageData.ContentType = imagePart.ContentType;
                    using var stream = imagePart.GetStream();
                    using var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    imageData.Data = ms.ToArray();
                }
            }
            catch
            {
                // Image extraction failed, continue without data
            }
        }

        // Get dimensions in EMUs for precise round-trip
        if (extent is not null)
        {
            imageData.WidthEmu = extent.Cx?.Value ?? 0;
            imageData.HeightEmu = extent.Cy?.Value ?? 0;
            imageData.WidthInches = imageData.WidthEmu / 914400.0;
            imageData.HeightInches = imageData.HeightEmu / 914400.0;
        }

        // Get alt text and description
        if (docPr is not null)
        {
            imageData.Name = docPr.Name?.Value ?? "";
            imageData.Description = docPr.Description?.Value;
            imageData.AltText = docPr.Title?.Value;
        }

        // Extract image formatting/positioning
        imageData.Formatting = ExtractImageFormatting(inline, anchor);

        var node = new DocumentNode(ContentType.Image, $"[Image: {imageData.Name}]");
        node.Metadata["ImageData"] = imageData;
        node.Metadata["Width"] = imageData.WidthInches;
        node.Metadata["Height"] = imageData.HeightInches;
        node.Metadata["ContentType"] = imageData.ContentType;

        return node;
    }

    /// <summary>
    /// Extracts image formatting and positioning
    /// </summary>
    private static ImageFormatting ExtractImageFormatting(DW.Inline? inline, DW.Anchor? anchor)
    {
        var formatting = new ImageFormatting();

        if (inline is not null)
        {
            formatting.IsInline = true;
            formatting.DistanceFromTop = inline.DistanceFromTop?.Value;
            formatting.DistanceFromBottom = inline.DistanceFromBottom?.Value;
            formatting.DistanceFromLeft = inline.DistanceFromLeft?.Value;
            formatting.DistanceFromRight = inline.DistanceFromRight?.Value;
        }
        else if (anchor is not null)
        {
            formatting.IsInline = false;
            formatting.DistanceFromTop = anchor.DistanceFromTop?.Value;
            formatting.DistanceFromBottom = anchor.DistanceFromBottom?.Value;
            formatting.DistanceFromLeft = anchor.DistanceFromLeft?.Value;
            formatting.DistanceFromRight = anchor.DistanceFromRight?.Value;
            formatting.AllowOverlap = anchor.AllowOverlap?.Value ?? false;
            formatting.BehindDocument = anchor.BehindDoc?.Value ?? false;
            formatting.LayoutInCell = anchor.LayoutInCell?.Value ?? false;
            formatting.Locked = anchor.Locked?.Value ?? false;
            formatting.RelativeHeight = anchor.RelativeHeight?.Value;

            // Horizontal position
            var hPos = anchor.HorizontalPosition;
            if (hPos is not null)
            {
                formatting.HorizontalRelativeTo = hPos.RelativeFrom?.Value.ToString();
                formatting.HorizontalPosition = hPos.PositionOffset?.Text;
            }

            // Vertical position
            var vPos = anchor.VerticalPosition;
            if (vPos is not null)
            {
                formatting.VerticalRelativeTo = vPos.RelativeFrom?.Value.ToString();
                formatting.VerticalPosition = vPos.PositionOffset?.Text;
            }

            // Wrap type
            formatting.WrapType = anchor switch
            {
                _ when anchor.GetFirstChild<DW.WrapNone>() is not null => "None",
                _ when anchor.GetFirstChild<DW.WrapSquare>() is not null => "Square",
                _ when anchor.GetFirstChild<DW.WrapTight>() is not null => "Tight",
                _ when anchor.GetFirstChild<DW.WrapThrough>() is not null => "Through",
                _ when anchor.GetFirstChild<DW.WrapTopBottom>() is not null => "TopAndBottom",
                _ => null
            };
        }

        return formatting;
    }

    /// <summary>
    /// Processes structured document tags (content controls)
    /// </summary>
    private DocumentNode? ProcessStructuredDocumentTag(SdtBlock sdtBlock)
    {
        var content = sdtBlock.SdtContentBlock;
        if (content is null) return null;

        // Get all child elements (paragraphs and tables)
        var childElements = content.ChildElements.ToList();

        // If there's only one paragraph, process it normally but preserve SDT context
        var paragraphs = content.Elements<Paragraph>().ToList();
        var tables = content.Elements<Table>().ToList();

        if (paragraphs.Count == 1 && tables.Count == 0)
        {
            var paraNode = ProcessParagraph(paragraphs[0]);
            if (paraNode is not null)
            {
                // Store the entire SDT block XML to preserve structure for complex content
                paraNode.OriginalXml = sdtBlock.OuterXml;
                paraNode.Metadata["IsSdtContent"] = true;
            }
            return paraNode;
        }

        // For SDT blocks with multiple elements, create a container node and store original XML
        var containerNode = new DocumentNode(ContentType.Paragraph, "")
        {
            OriginalXml = sdtBlock.OuterXml,
            Metadata = { ["IsSdtBlock"] = true }
        };

        // Process each child element
        foreach (var element in childElements)
        {
            var childNode = element switch
            {
                Paragraph para => ProcessParagraph(para),
                Table table => ProcessTable(table),
                _ => null
            };

            if (childNode is not null)
            {
                containerNode.AddChild(childNode);
            }
        }

        // Return the container if it has children, otherwise null
        return containerNode.Children.Count > 0 ? containerNode : null;
    }

    public void Dispose()
    {
        _document?.Dispose();
        _document = null;
        _mainPart = null;
    }
}
