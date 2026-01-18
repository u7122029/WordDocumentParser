using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordDocumentParser
{
    /// <summary>
    /// Writes a document tree structure to a Word document (.docx file).
    /// Preserves all formatting for round-trip fidelity by restoring original document parts.
    /// </summary>
    public class WordDocumentTreeWriter : IDisposable
    {
        private WordprocessingDocument? _document;
        private MainDocumentPart? _mainPart;
        private Body? _body;
        private NumberingDefinitionsPart? _numberingPart;
        private int _currentListId = 1;
        private uint _imageCounter = 1;
        private readonly Dictionary<string, string> _hyperlinkRelationships = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _imageRelationshipMapping = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _hyperlinkRelationshipMapping = new Dictionary<string, string>();
        private DocumentPackageData? _packageData;

        /// <summary>
        /// Writes a document tree to a file
        /// </summary>
        public void WriteToFile(DocumentNode root, string filePath)
        {
            _document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            try
            {
                BuildDocument(root);
                _document.Save();
            }
            finally
            {
                _document.Dispose();
                _document = null;
            }
        }

        /// <summary>
        /// Writes a document tree to a stream
        /// </summary>
        public void WriteToStream(DocumentNode root, Stream stream)
        {
            _document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, false);
            try
            {
                BuildDocument(root);
                _document.Save();
            }
            finally
            {
                _document.Dispose();
                _document = null;
            }
        }

        /// <summary>
        /// Builds the document content from the tree
        /// </summary>
        private void BuildDocument(DocumentNode root)
        {
            _packageData = root.PackageData;
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
            ProcessNode(root);

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
            // Restore styles
            if (!string.IsNullOrEmpty(_packageData!.StylesXml))
            {
                var stylesPart = _mainPart!.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(_packageData.StylesXml);
            }
            else
            {
                AddStyleDefinitions();
            }

            // Restore theme
            if (!string.IsNullOrEmpty(_packageData.ThemeXml))
            {
                var themePart = _mainPart!.AddNewPart<ThemePart>();
                using var reader = new StringReader(_packageData.ThemeXml);
                themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(_packageData.ThemeXml);
            }

            // Restore font table
            if (!string.IsNullOrEmpty(_packageData.FontTableXml))
            {
                var fontTablePart = _mainPart!.AddNewPart<FontTablePart>();
                fontTablePart.Fonts = new Fonts(_packageData.FontTableXml);
            }

            // Restore numbering definitions
            if (!string.IsNullOrEmpty(_packageData.NumberingXml))
            {
                _numberingPart = _mainPart!.AddNewPart<NumberingDefinitionsPart>();
                _numberingPart.Numbering = new Numbering(_packageData.NumberingXml);
            }

            // Restore document settings
            if (!string.IsNullOrEmpty(_packageData.SettingsXml))
            {
                var settingsPart = _mainPart!.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(_packageData.SettingsXml);
            }

            // Restore web settings
            if (!string.IsNullOrEmpty(_packageData.WebSettingsXml))
            {
                var webSettingsPart = _mainPart!.AddNewPart<WebSettingsPart>();
                webSettingsPart.WebSettings = new WebSettings(_packageData.WebSettingsXml);
            }

            // Restore footnotes
            if (!string.IsNullOrEmpty(_packageData.FootnotesXml))
            {
                var footnotesPart = _mainPart!.AddNewPart<FootnotesPart>();
                footnotesPart.Footnotes = new Footnotes(_packageData.FootnotesXml);
            }

            // Restore endnotes
            if (!string.IsNullOrEmpty(_packageData.EndnotesXml))
            {
                var endnotesPart = _mainPart!.AddNewPart<EndnotesPart>();
                endnotesPart.Endnotes = new Endnotes(_packageData.EndnotesXml);
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

            // Restore headers
            foreach (var kvp in _packageData.Headers)
            {
                var headerPart = _mainPart!.AddNewPart<HeaderPart>();
                var headerXml = UpdateImageRelationships(kvp.Value);
                headerPart.Header = new Header(headerXml);
            }

            // Restore footers
            foreach (var kvp in _packageData.Footers)
            {
                var footerPart = _mainPart!.AddNewPart<FooterPart>();
                var footerXml = UpdateImageRelationships(kvp.Value);
                footerPart.Footer = new Footer(footerXml);
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
        }

        /// <summary>
        /// Updates image relationship IDs in XML content to use new relationship IDs
        /// </summary>
        private string UpdateImageRelationships(string xml)
        {
            if (_imageRelationshipMapping.Count == 0)
                return xml;

            var result = xml;
            foreach (var kvp in _imageRelationshipMapping)
            {
                // Replace relationship IDs in embed attributes
                result = result.Replace($"r:embed=\"{kvp.Key}\"", $"r:embed=\"{kvp.Value}\"");
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

                if (!string.IsNullOrEmpty(core.Title)) props.Title = core.Title;
                if (!string.IsNullOrEmpty(core.Subject)) props.Subject = core.Subject;
                if (!string.IsNullOrEmpty(core.Creator)) props.Creator = core.Creator;
                if (!string.IsNullOrEmpty(core.Keywords)) props.Keywords = core.Keywords;
                if (!string.IsNullOrEmpty(core.Description)) props.Description = core.Description;
                if (!string.IsNullOrEmpty(core.LastModifiedBy)) props.LastModifiedBy = core.LastModifiedBy;
                if (!string.IsNullOrEmpty(core.Revision)) props.Revision = core.Revision;
                if (!string.IsNullOrEmpty(core.Category)) props.Category = core.Category;
                if (!string.IsNullOrEmpty(core.ContentStatus)) props.ContentStatus = core.ContentStatus;

                if (!string.IsNullOrEmpty(core.Created) && DateTime.TryParse(core.Created, out var created))
                    props.Created = created;
                if (!string.IsNullOrEmpty(core.Modified) && DateTime.TryParse(core.Modified, out var modified))
                    props.Modified = modified;
            }

            // Restore extended properties
            if (_packageData?.ExtendedProperties != null)
            {
                var extPart = _document!.AddExtendedFilePropertiesPart();
                var extProps = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                var ext = _packageData.ExtendedProperties;

                if (!string.IsNullOrEmpty(ext.Template))
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Template(ext.Template));
                if (!string.IsNullOrEmpty(ext.Application))
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Application(ext.Application));
                if (!string.IsNullOrEmpty(ext.AppVersion))
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.ApplicationVersion(ext.AppVersion));
                if (!string.IsNullOrEmpty(ext.Company))
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Company(ext.Company));
                if (!string.IsNullOrEmpty(ext.Manager))
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Manager(ext.Manager));
                if (ext.Pages.HasValue)
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Pages(ext.Pages.Value.ToString()));
                if (ext.Words.HasValue)
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Words(ext.Words.Value.ToString()));
                if (ext.Characters.HasValue)
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Characters(ext.Characters.Value.ToString()));
                if (ext.CharactersWithSpaces.HasValue)
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.CharactersWithSpaces(ext.CharactersWithSpaces.Value.ToString()));
                if (ext.Lines.HasValue)
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Lines(ext.Lines.Value.ToString()));
                if (ext.Paragraphs.HasValue)
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.Paragraphs(ext.Paragraphs.Value.ToString()));
                if (ext.TotalTime.HasValue)
                    extProps.Append(new DocumentFormat.OpenXml.ExtendedProperties.TotalTime(ext.TotalTime.Value.ToString()));

                extPart.Properties = extProps;
            }

            // Restore custom properties
            if (!string.IsNullOrEmpty(_packageData?.CustomPropertiesXml))
            {
                var customPart = _document!.AddCustomFilePropertiesPart();
                customPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(_packageData.CustomPropertiesXml);
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
            }
        }

        /// <summary>
        /// Writes a heading paragraph with full formatting
        /// </summary>
        private void WriteHeading(DocumentNode node)
        {
            // Use original XML if available for exact round-trip
            if (!string.IsNullOrEmpty(node.OriginalXml))
            {
                var updatedXml = UpdateImageRelationships(node.OriginalXml);
                var paragraph = new Paragraph(updatedXml);
                _body!.Append(paragraph);
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
            foreach (var child in node.Children)
            {
                ProcessNode(child);
            }
        }

        /// <summary>
        /// Writes a regular paragraph with full formatting
        /// </summary>
        private void WriteParagraph(DocumentNode node)
        {
            // Use original XML if available for exact round-trip
            if (!string.IsNullOrEmpty(node.OriginalXml))
            {
                var updatedXml = UpdateImageRelationships(node.OriginalXml);
                var paragraph = new Paragraph(updatedXml);
                _body!.Append(paragraph);
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

            // Process other children
            foreach (var child in node.Children)
            {
                if (child.Type != ContentType.Image)
                {
                    ProcessNode(child);
                }
            }
        }

        /// <summary>
        /// Creates paragraph properties from formatting
        /// </summary>
        private ParagraphProperties CreateParagraphProperties(DocumentNode node)
        {
            var props = new ParagraphProperties();
            var fmt = node.ParagraphFormatting;

            if (fmt == null) return props;

            // Style
            if (!string.IsNullOrEmpty(fmt.StyleId))
            {
                props.Append(new ParagraphStyleId { Val = fmt.StyleId });
            }

            // Alignment
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

            // Indentation
            if (fmt.IndentLeft != null || fmt.IndentRight != null || fmt.IndentFirstLine != null || fmt.IndentHanging != null)
            {
                var ind = new Indentation();
                if (fmt.IndentLeft != null) ind.Left = fmt.IndentLeft;
                if (fmt.IndentRight != null) ind.Right = fmt.IndentRight;
                if (fmt.IndentFirstLine != null) ind.FirstLine = fmt.IndentFirstLine;
                if (fmt.IndentHanging != null) ind.Hanging = fmt.IndentHanging;
                props.Append(ind);
            }

            // Spacing
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

            // Keep with next/keep lines
            if (fmt.KeepNext) props.Append(new KeepNext());
            if (fmt.KeepLines) props.Append(new KeepLines());
            if (fmt.PageBreakBefore) props.Append(new PageBreakBefore());
            if (fmt.WidowControl) props.Append(new WidowControl());

            // Shading
            if (!string.IsNullOrEmpty(fmt.ShadingFill))
            {
                props.Append(new Shading { Fill = fmt.ShadingFill, Color = fmt.ShadingColor });
            }

            // Borders
            if (fmt.TopBorder != null || fmt.BottomBorder != null || fmt.LeftBorder != null || fmt.RightBorder != null)
            {
                var borders = new ParagraphBorders();
                if (fmt.TopBorder != null) borders.Append(CreateBorder<TopBorder>(fmt.TopBorder));
                if (fmt.BottomBorder != null) borders.Append(CreateBorder<BottomBorder>(fmt.BottomBorder));
                if (fmt.LeftBorder != null) borders.Append(CreateBorder<LeftBorder>(fmt.LeftBorder));
                if (fmt.RightBorder != null) borders.Append(CreateBorder<RightBorder>(fmt.RightBorder));
                props.Append(borders);
            }

            // Numbering (for list items)
            if (fmt.NumberingId.HasValue)
            {
                props.Append(new NumberingProperties(
                    new NumberingLevelReference { Val = fmt.NumberingLevel ?? 0 },
                    new NumberingId { Val = fmt.NumberingId.Value }
                ));
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
        /// Creates run properties from formatting
        /// </summary>
        private RunProperties CreateRunProperties(RunFormatting fmt)
        {
            var props = new RunProperties();

            // Style
            if (!string.IsNullOrEmpty(fmt.StyleId))
            {
                props.Append(new RunStyle { Val = fmt.StyleId });
            }

            // Bold
            if (fmt.Bold)
            {
                props.Append(new Bold());
            }

            // Italic
            if (fmt.Italic)
            {
                props.Append(new Italic());
            }

            // Underline
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

            // Strike
            if (fmt.Strike)
            {
                props.Append(new Strike());
            }
            if (fmt.DoubleStrike)
            {
                props.Append(new DoubleStrike());
            }

            // Font
            if (fmt.FontFamily != null || fmt.FontFamilyAscii != null)
            {
                var fonts = new RunFonts();
                if (fmt.FontFamily != null) fonts.HighAnsi = fmt.FontFamily;
                if (fmt.FontFamilyAscii != null) fonts.Ascii = fmt.FontFamilyAscii;
                if (fmt.FontFamilyEastAsia != null) fonts.EastAsia = fmt.FontFamilyEastAsia;
                if (fmt.FontFamilyComplexScript != null) fonts.ComplexScript = fmt.FontFamilyComplexScript;
                props.Append(fonts);
            }

            // Font size
            if (!string.IsNullOrEmpty(fmt.FontSize))
            {
                props.Append(new FontSize { Val = fmt.FontSize });
            }
            if (!string.IsNullOrEmpty(fmt.FontSizeComplexScript))
            {
                props.Append(new FontSizeComplexScript { Val = fmt.FontSizeComplexScript });
            }

            // Color
            if (!string.IsNullOrEmpty(fmt.Color))
            {
                props.Append(new Color { Val = fmt.Color });
            }

            // Highlight
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

            // Superscript/Subscript
            if (fmt.Superscript)
            {
                props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
            }
            else if (fmt.Subscript)
            {
                props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
            }

            // Caps
            if (fmt.SmallCaps)
            {
                props.Append(new SmallCaps());
            }
            if (fmt.AllCaps)
            {
                props.Append(new Caps());
            }

            // Shading
            if (!string.IsNullOrEmpty(fmt.Shading))
            {
                props.Append(new Shading { Fill = fmt.Shading });
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
                var updatedXml = UpdateImageRelationships(node.OriginalXml);
                var originalTable = new Table(updatedXml);
                _body!.Append(originalTable);
                _body.Append(new Paragraph()); // spacing after table
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
        /// Creates table cell properties from formatting
        /// </summary>
        private TableCellProperties CreateTableCellProperties(TableCellFormatting? fmt)
        {
            var props = new TableCellProperties();

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
                    props.Append(new TableCellWidth { Width = fmt.Width, Type = widthType });
                }

                // Grid span
                if (fmt.GridSpan > 1)
                {
                    props.Append(new GridSpan { Val = fmt.GridSpan });
                }

                // Vertical merge
                if (!string.IsNullOrEmpty(fmt.VerticalMerge))
                {
                    var vMerge = new VerticalMerge();
                    if (fmt.VerticalMerge == "Restart")
                    {
                        vMerge.Val = MergedCellValues.Restart;
                    }
                    props.Append(vMerge);
                }

                // Vertical alignment
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

                // Shading
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

                // Borders
                if (fmt.TopBorder != null || fmt.BottomBorder != null || fmt.LeftBorder != null || fmt.RightBorder != null)
                {
                    var borders = new TableCellBorders();
                    if (fmt.TopBorder != null) borders.Append(CreateBorder<TopBorder>(fmt.TopBorder));
                    if (fmt.BottomBorder != null) borders.Append(CreateBorder<BottomBorder>(fmt.BottomBorder));
                    if (fmt.LeftBorder != null) borders.Append(CreateBorder<LeftBorder>(fmt.LeftBorder));
                    if (fmt.RightBorder != null) borders.Append(CreateBorder<RightBorder>(fmt.RightBorder));
                    props.Append(borders);
                }

                // No wrap
                if (fmt.NoWrap)
                {
                    props.Append(new NoWrap());
                }
            }

            return props;
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
                var sectionProps = new SectionProperties(lastSectPrXml);

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
            // Get the header and footer parts with their new relationship IDs
            var headerParts = _mainPart?.HeaderParts?.ToList() ?? new List<HeaderPart>();
            var footerParts = _mainPart?.FooterParts?.ToList() ?? new List<FooterPart>();

            // Update header references
            var headerRefs = sectionProps.Elements<HeaderReference>().ToList();
            for (int i = 0; i < headerRefs.Count && i < headerParts.Count; i++)
            {
                headerRefs[i].Id = _mainPart!.GetIdOfPart(headerParts[i]);
            }

            // Update footer references
            var footerRefs = sectionProps.Elements<FooterReference>().ToList();
            for (int i = 0; i < footerRefs.Count && i < footerParts.Count; i++)
            {
                footerRefs[i].Id = _mainPart!.GetIdOfPart(footerParts[i]);
            }
        }

        public void Dispose()
        {
            _document?.Dispose();
            _document = null;
            _mainPart = null;
            _body = null;
        }
    }
}
