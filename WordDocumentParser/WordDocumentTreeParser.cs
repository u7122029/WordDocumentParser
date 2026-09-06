using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Models.ContentControls;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Package;
using WordDocumentParser.Parsing;
using WordDocumentParser.Parsing.Extractors;

namespace WordDocumentParser;

/// <summary>
/// Parses Word documents and builds a hierarchical tree structure based on headings.
/// Captures full formatting and document package data for round-trip fidelity.
/// </summary>
public class WordDocumentTreeParser : IDocumentParser
{
    private WordprocessingDocument? _document;
    private MainDocumentPart? _mainPart;
    private ParsingContext? _context;
    private ImageExtractor? _imageExtractor;
    private TableExtractor? _tableExtractor;
    private byte[]? _packageBytes;

    /// <summary>
    /// Bounds applied while parsing, so a hostile document cannot exhaust memory.
    /// Defaults to <see cref="DocumentLimits.Trusted"/>, which applies no bounds; pass
    /// <see cref="DocumentLimits.Untrusted"/> when parsing documents from outside the trust boundary.
    /// </summary>
    public DocumentLimits Limits { get; init; } = DocumentLimits.Trusted;

    /// <summary>
    /// What to do when part of the source document cannot be preserved. By default such failures
    /// throw, so a lossy parse is never reported as a successful one.
    /// </summary>
    public RecoveryOptions Recovery { get; init; } = new();

    /// <summary>
    /// Whether to keep a copy of the source package so the writer can edit it in place.
    /// </summary>
    /// <remarks>
    /// On by default, and required for faithful round-trips: it is what lets parts this library has
    /// no model for pass through untouched. Turning it off avoids holding a copy of the source
    /// package, at the cost of reassembling the package from the model on save — which loses
    /// anything the model does not represent.
    /// </remarks>
    public bool CaptureOriginalPackage { get; init; } = true;

    /// <summary>
    /// Whether to keep the full document XML in <see cref="DocumentPackageData.OriginalDocumentXml"/>.
    /// </summary>
    /// <remarks>
    /// Off by default. Each node already holds its own XML and, when
    /// <see cref="CaptureOriginalPackage"/> is on, the whole package is retained too — so this was a
    /// third redundant copy of the document's text.
    /// </remarks>
    public bool CaptureOriginalDocumentXml { get; init; }

    /// <summary>
    /// Parses a Word document from a file path and returns the document
    /// </summary>
    public WordDocument ParseFromFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Document not found: {filePath}");

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return ParseFromStream(stream, Path.GetFileName(filePath));
    }

    /// <summary>
    /// Parses a Word document from a stream and returns the document.
    /// </summary>
    /// <param name="stream">The <c>.docx</c> package to read. The stream is read but not disposed.</param>
    /// <param name="documentName">A name recorded on the returned document.</param>
    /// <returns>The parsed document.</returns>
    /// <remarks>
    /// A parser instance may be reused sequentially — each call releases the previous document's
    /// package first — but not concurrently. Use one instance per thread.
    /// </remarks>
    public WordDocument ParseFromStream(Stream stream, string documentName = "Document")
    {
        ArgumentNullException.ThrowIfNull(stream);

        // Release anything held from a previous parse before taking new state, so reusing an
        // instance cannot leak the old package or mix its parts into the new document.
        ReleaseParsingState();

        _packageBytes = CaptureOriginalPackage ? ReadPackageBytes(stream) : null;

        var source = _packageBytes is null ? stream : new MemoryStream(_packageBytes, writable: false);
        try
        {
            _document = WordprocessingDocument.Open(source, false, new OpenSettings
            {
                MaxCharactersInPart = Limits.MaxCharactersInPart
            });
            _mainPart = _document.MainDocumentPart;

            if (_mainPart is null)
                throw new InvalidOperationException("Invalid Word document: missing main document part");

            // Both checks must precede any DOM access. The SDK builds its object model by recursive
            // descent, so a deeply nested document overflows the stack inside the SDK — a crash no
            // caller can catch — before a limit checked during our own traversal would ever run.
            Limits.EnforcePartCount(CountParts(_document));
            EnforceNestingLimit(_document);

            if (_mainPart.Document?.Body is null)
                throw new InvalidOperationException("Invalid Word document: missing body");

            _context = new ParsingContext
            {
                Document = _document,
                MainPart = _mainPart,
                Limits = Limits,
                Recovery = Recovery
            };
            _context.CacheStyles();
            _context.CacheHyperlinkRelationships();
            _imageExtractor = new ImageExtractor(_context);
            _tableExtractor = new TableExtractor(_context, ProcessParagraph, ProcessTable);

            // Create root node
            var root = new DocumentNode(ContentType.Document, documentName)
            {
                Metadata = { ["FileName"] = documentName }
            };

            // Extract document package data for round-trip fidelity
            var packageData = ExtractPackageData();

            // Parse all body elements
            var bodyElements = _mainPart.Document.Body.Elements().ToList();
            BuildTree(root, bodyElements);

            // Everything now mirrors the source exactly. Recording that as the baseline lets the
            // writer distinguish values it merely parsed from values a caller later assigned.
            root.AcceptAllChanges();
            packageData.CaptureBaseline();

            var result = new WordDocument(root, packageData)
            {
                FileName = documentName
            };

            // The SDK DOM and its buffers are no longer needed: the tree holds each node's XML and
            // the package data holds every part. Releasing here keeps a parsed document's footprint
            // to the model rather than the model plus a live package.
            ReleaseParsingState();

            return result;
        }
        finally
        {
            if (!ReferenceEquals(source, stream))
            {
                source.Dispose();
            }
        }
    }

    /// <summary>
    /// Reads the source package into memory, subject to the binary budget.
    /// </summary>
    /// <remarks>
    /// The compressed package counts against the same budget as the media inside it: an oversized
    /// upload should be rejected before it is buffered, not after.
    /// </remarks>
    private byte[] ReadPackageBytes(Stream stream)
    {
        if (Limits.MaxTotalBinaryBytes > 0 && stream.CanSeek)
        {
            Limits.EnforceBinaryBudget(stream.Length - stream.Position);
        }

        if (stream is MemoryStream alreadyBuffered && alreadyBuffered.Position == 0)
        {
            return alreadyBuffered.ToArray();
        }

        return ParsingContext.ReadBounded(stream, Limits, alreadyRead: 0, target: "the document package");
    }

    /// <summary>
    /// Streams every XML part and rejects the document if any nests deeper than allowed.
    /// </summary>
    /// <remarks>
    /// <see cref="XmlReader"/> is iterative, so this scan cannot itself overflow the stack, and it
    /// runs before the SDK is asked for an object model. The SDK's own loader descends recursively:
    /// a two-kilobyte package of ten thousand nested tags takes the process down with it, and a
    /// stack overflow cannot be caught. Skipped entirely when no depth limit is configured.
    /// </remarks>
    private void EnforceNestingLimit(WordprocessingDocument document)
    {
        if (Limits.MaxElementDepth <= 0) return;

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            IgnoreProcessingInstructions = true,
            CloseInput = false
        };

        foreach (var part in EnumerateParts(document))
        {
            if (!part.ContentType.Contains("xml", StringComparison.OrdinalIgnoreCase)) continue;

            try
            {
                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, settings);

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        // Depth is zero-based at the root element.
                        Limits.EnforceDepth(reader.Depth + 1);
                    }
                }
            }
            catch (XmlException ex)
            {
                throw new DocumentPreservationException(
                    part.Uri.ToString(), "Part is not well-formed XML.", ex);
            }
        }
    }

    /// <summary>
    /// Walks every part reachable from the package, without loading any object model.
    /// </summary>
    private static IEnumerable<OpenXmlPart> EnumerateParts(WordprocessingDocument document)
    {
        var seen = new HashSet<Uri>();
        var pending = new Queue<OpenXmlPartContainer>();
        pending.Enqueue(document);

        while (pending.Count > 0)
        {
            foreach (var pair in pending.Dequeue().Parts)
            {
                if (!seen.Add(pair.OpenXmlPart.Uri)) continue;

                pending.Enqueue(pair.OpenXmlPart);
                yield return pair.OpenXmlPart;
            }
        }
    }

    private static int CountParts(WordprocessingDocument document)
    {
        var seen = new HashSet<Uri>();
        var pending = new Queue<OpenXmlPartContainer>();
        pending.Enqueue(document);

        while (pending.Count > 0)
        {
            foreach (var pair in pending.Dequeue().Parts)
            {
                if (seen.Add(pair.OpenXmlPart.Uri))
                {
                    pending.Enqueue(pair.OpenXmlPart);
                }
            }
        }

        return seen.Count;
    }

    /// <summary>
    /// Disposes the open package and clears every field tied to it.
    /// </summary>
    private void ReleaseParsingState()
    {
        _document?.Dispose();
        _document = null;
        _mainPart = null;
        _context = null;
        _imageExtractor = null;
        _tableExtractor = null;
    }

    /// <summary>
    /// Extracts all document package data for round-trip preservation
    /// </summary>
    private DocumentPackageData ExtractPackageData()
    {
        var packageData = new DocumentPackageData
        {
            // Keeping the source package is what makes preservation total: on save the writer edits
            // a copy of it, so comments, revisions, embedded objects, the relationships owned by
            // headers and footers, and extension-namespace elements all survive without this
            // library needing a model for any of them.
            OriginalPackageBytes = _packageBytes
        };

        // Store original document XML
        if (CaptureOriginalDocumentXml && _mainPart?.Document is not null)
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

        // Extract headers and their images
        foreach (var headerPart in _mainPart?.HeaderParts ?? [])
        {
            var relId = _mainPart!.GetIdOfPart(headerPart);
            if (headerPart.Header is not null)
            {
                packageData.Headers[relId] = headerPart.Header.OuterXml;
            }

            // Extract images from this header
            var headerImages = ReadImageParts(headerPart, headerPart.ImageParts);
            if (headerImages.Count > 0)
            {
                packageData.HeaderImages[relId] = headerImages;
            }
        }

        // Extract footers and their images
        foreach (var footerPart in _mainPart?.FooterParts ?? [])
        {
            var relId = _mainPart!.GetIdOfPart(footerPart);
            if (footerPart.Footer is not null)
            {
                packageData.Footers[relId] = footerPart.Footer.OuterXml;
            }

            // Extract images from this footer
            var footerImages = ReadImageParts(footerPart, footerPart.ImageParts);
            if (footerImages.Count > 0)
            {
                packageData.FooterImages[relId] = footerImages;
            }
        }

        // Extract images
        if (_mainPart is not null)
        {
            packageData.Images = ReadImageParts(_mainPart, _mainPart.ImageParts);
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

        // Record every relationship ID the main part uses, so merging allocates against all of them
        // rather than colliding with a kind of resource it did not happen to look at.
        if (_mainPart is not null)
        {
            foreach (var pair in _mainPart.Parts)
            {
                packageData.KnownRelationshipIds.Add(pair.RelationshipId);
            }

            foreach (var rel in _mainPart.HyperlinkRelationships)
            {
                packageData.KnownRelationshipIds.Add(rel.Id);
            }

            foreach (var rel in _mainPart.ExternalRelationships)
            {
                packageData.KnownRelationshipIds.Add(rel.Id);
            }
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
            packageData.CorePropertiesXml = _context!.ReadTextPart(_document.CoreFilePropertiesPart);
        }

        // Extract raw app properties XML for exact round-trip
        if (_document?.ExtendedFilePropertiesPart is not null)
        {
            packageData.AppPropertiesXml = _context!.ReadTextPart(_document.ExtendedFilePropertiesPart);
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
                // Read through the bounded reader: these parts never pass through the SDK's parser,
                // so its own per-part character cap does not apply to them.
                var xmlContent = _context!.ReadTextPart(customXmlPart);

                string? propsXml = null;
                if (customXmlPart.CustomXmlPropertiesPart is not null)
                {
                    propsXml = _context.ReadTextPart(customXmlPart.CustomXmlPropertiesPart);
                }

                packageData.CustomXmlParts[customXmlPart.Uri.ToString()] = new CustomXmlPartData
                {
                    XmlContent = xmlContent,
                    PropertiesXml = propsXml
                };
            }
            catch (Exception ex) when (ex is not DocumentLimitExceededException)
            {
                Recovery.Report(customXmlPart.Uri.ToString(), "Custom XML part could not be read.", ex);
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
            packageData.GlossaryImages = ReadImageParts(glossaryPart, glossaryPart.ImageParts);
        }

        return packageData;
    }

    /// <summary>
    /// Reads a container's image parts into shared buffers keyed by relationship ID.
    /// </summary>
    private Dictionary<string, ImagePartData> ReadImageParts(
        OpenXmlPartContainer owner,
        IEnumerable<ImagePart> imageParts)
    {
        var images = new Dictionary<string, ImagePartData>();

        foreach (var imagePart in imageParts)
        {
            var relId = owner.GetIdOfPart(imagePart);
            try
            {
                images[relId] = new ImagePartData
                {
                    ContentType = imagePart.ContentType,
                    Data = _context!.ReadBinaryPart(imagePart),
                    OriginalRelationshipId = relId,
                    OriginalUri = imagePart.Uri?.ToString()
                };
            }
            catch (Exception ex) when (ex is not DocumentLimitExceededException)
            {
                Recovery.Report(relId, "Image part could not be read.", ex);
            }
        }

        return images;
    }

    /// <summary>
    /// Builds the tree structure by processing body elements and organizing by heading hierarchy.
    /// </summary>
    /// <remarks>
    /// Section membership is a stack discipline: each heading closes every open heading at its level
    /// or deeper before it is placed. Leaving deeper entries open would let a later, deeper heading
    /// attach beneath a stale ancestor — landing in the wrong section and out of document order.
    /// </remarks>
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

                // Close the sections this heading ends: everything at its level or deeper.
                for (var deeper = level; deeper <= maxLevel; deeper++)
                {
                    stack[deeper] = null;
                }

                // Attach to the nearest still-open shallower heading.
                var parentLevel = level - 1;
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

        // For round-trip fidelity, preserve ALL paragraphs including empty ones.
        // Empty paragraphs serve important purposes: spacing, formatting, structure.
        // We store OriginalXml for every paragraph to ensure exact reproduction.

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
        node.ParagraphFormatting = _context is null
            ? new ParagraphFormatting()
            : FormattingExtractor.ExtractParagraphFormatting(para, _context);

        // Check for images in the paragraph
        var images = _imageExtractor?.ExtractImages(para) ?? [];
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

        // The node now mirrors the source paragraph exactly, so this is its baseline.
        node.AcceptAllChanges();

        return node;
    }

    /// <summary>
    /// Extracts formatted runs from a paragraph, preserving all text formatting
    /// </summary>
    private List<FormattedRun> ExtractFormattedRuns(Paragraph para)
    {
        List<FormattedRun> formattedRuns = [];

        // Track field state for complex fields (fldChar-based fields like DOCPROPERTY)
        var inField = false;
        string? currentFieldCode = null;
        var fieldResultRuns = new List<FormattedRun>();

        foreach (var child in para.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    // Check for field characters
                    var fldChar = run.GetFirstChild<FieldChar>();
                    if (fldChar?.FieldCharType?.Value is not null)
                    {
                        var charType = fldChar.FieldCharType.Value;
                        if (charType == FieldCharValues.Begin)
                        {
                            inField = true;
                            currentFieldCode = null;
                            fieldResultRuns.Clear();
                        }
                        else if (charType == FieldCharValues.Separate)
                        {
                            // Field code is complete, now collecting result
                        }
                        else if (charType == FieldCharValues.End)
                        {
                            // Process the completed field
                            if (currentFieldCode is not null && IsDocPropertyField(currentFieldCode))
                            {
                                var propInfo = ParseDocPropertyField(currentFieldCode);
                                // Create a single run for the document property field
                                var fieldValue = string.Concat(fieldResultRuns.Select(r => r.Text));
                                var propField = new DocumentPropertyField
                                {
                                    PropertyName = propInfo.propertyName,
                                    PropertyType = propInfo.propertyType,
                                    Value = fieldValue,
                                    FieldCode = currentFieldCode
                                };

                                // Look up the actual value from document properties if available
                                propField.Value = ResolveDocumentPropertyValue(propInfo.propertyName, propInfo.propertyType) ?? fieldValue;

                                formattedRuns.Add(new FormattedRun
                                {
                                    Text = propField.Value ?? "",
                                    DocumentPropertyField = propField,
                                    Formatting = fieldResultRuns.FirstOrDefault()?.Formatting ?? new RunFormatting()
                                });
                            }
                            else
                            {
                                // Not a DOCPROPERTY field, add result runs normally
                                formattedRuns.AddRange(fieldResultRuns);
                            }
                            inField = false;
                            currentFieldCode = null;
                            fieldResultRuns.Clear();
                        }
                        continue;
                    }

                    // Check for field code
                    var fieldCode = run.GetFirstChild<FieldCode>();
                    if (fieldCode is not null && inField)
                    {
                        currentFieldCode = (currentFieldCode ?? "") + fieldCode.Text;
                        continue;
                    }

                    // Process the run
                    var runs = ProcessRun(run);
                    if (inField && currentFieldCode is not null)
                    {
                        // This is a field result run
                        fieldResultRuns.AddRange(runs);
                    }
                    else
                    {
                        formattedRuns.AddRange(runs);
                    }
                    break;

                case Hyperlink hyperlink:
                    formattedRuns.AddRange(ProcessHyperlinkRuns(hyperlink));
                    break;

                case SdtRun sdtRun:
                    // Handle inline content controls
                    formattedRuns.AddRange(ProcessInlineSdt(sdtRun));
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
    /// Checks if a field code is a DOCPROPERTY field
    /// </summary>
    private static bool IsDocPropertyField(string fieldCode)
    {
        var trimmed = fieldCode.Trim();
        return trimmed.StartsWith("DOCPROPERTY", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith(" DOCPROPERTY", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Parses a DOCPROPERTY field code to extract property information
    /// </summary>
    private static (string propertyName, DocumentPropertyType propertyType) ParseDocPropertyField(string fieldCode)
    {
        // Field code format: " DOCPROPERTY  PropertyName  \* MERGEFORMAT "
        var trimmed = fieldCode.Trim();
        var parts = trimmed.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);

        var propertyName = "";
        if (parts.Length >= 2)
        {
            // Skip "DOCPROPERTY" and get the property name
            // Handle quoted property names
            var nameIndex = 1;
            if (parts[nameIndex].StartsWith('"'))
            {
                // Find the closing quote
                var nameParts = new List<string>();
                for (var i = nameIndex; i < parts.Length; i++)
                {
                    nameParts.Add(parts[i]);
                    if (parts[i].EndsWith('"'))
                        break;
                }
                propertyName = string.Join(" ", nameParts).Trim('"');
            }
            else
            {
                propertyName = parts[nameIndex];
            }
        }

        return (propertyName, DocumentPropertyHelpers.DeterminePropertyType(propertyName));
    }

    /// <summary>
    /// Resolves a document property value from the cached document properties
    /// </summary>
    private string? ResolveDocumentPropertyValue(string propertyName, DocumentPropertyType propertyType)
    {
        if (_document is null) return null;

        return propertyType switch
        {
            DocumentPropertyType.Core => GetCorePropertyValue(propertyName),
            DocumentPropertyType.Extended => GetExtendedPropertyValue(propertyName),
            DocumentPropertyType.Custom => GetCustomPropertyValue(propertyName),
            _ => null
        };
    }

    /// <summary>
    /// Gets a core document property value
    /// </summary>
    private string? GetCorePropertyValue(string propertyName)
    {
        var props = _document?.PackageProperties;
        if (props is null) return null;

        return propertyName.ToLowerInvariant() switch
        {
            "title" => props.Title,
            "subject" => props.Subject,
            "creator" or "author" => props.Creator,
            "keywords" => props.Keywords,
            "description" or "comments" => props.Description,
            "lastmodifiedby" => props.LastModifiedBy,
            "revision" => props.Revision,
            "created" => props.Created?.ToString("g"),
            "modified" => props.Modified?.ToString("g"),
            "category" => props.Category,
            "contentstatus" or "status" => props.ContentStatus,
            _ => null
        };
    }

    /// <summary>
    /// Gets an extended document property value
    /// </summary>
    private string? GetExtendedPropertyValue(string propertyName)
    {
        var extProps = _document?.ExtendedFilePropertiesPart?.Properties;
        if (extProps is null) return null;

        return propertyName.ToLowerInvariant() switch
        {
            "template" => extProps.Template?.Text,
            "application" => extProps.Application?.Text,
            "appversion" => extProps.ApplicationVersion?.Text,
            "company" => extProps.Company?.Text,
            "manager" => extProps.Manager?.Text,
            "pages" => extProps.Pages?.Text,
            "words" => extProps.Words?.Text,
            "characters" => extProps.Characters?.Text,
            "characterswithspaces" => extProps.CharactersWithSpaces?.Text,
            "lines" => extProps.Lines?.Text,
            "paragraphs" => extProps.Paragraphs?.Text,
            "totaltime" => extProps.TotalTime?.Text,
            _ => null
        };
    }

    /// <summary>
    /// Gets a custom document property value
    /// </summary>
    private string? GetCustomPropertyValue(string propertyName)
    {
        var customProps = _document?.CustomFilePropertiesPart?.Properties;
        if (customProps is null) return null;

        foreach (var prop in customProps.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
        {
            if (string.Equals(prop.Name?.Value, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                // Get the value from the property (could be various types)
                return prop.InnerText;
            }
        }

        return null;
    }

    /// <summary>
    /// Processes a single run element and extracts formatted runs
    /// </summary>
    private List<FormattedRun> ProcessRun(Run run)
    {
        List<FormattedRun> result = [];
        var formatting = FormattingExtractor.ExtractRunFormatting(run.RunProperties);

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
            data.Url = _context?.GetHyperlinkUrl(data.RelationshipId);

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

        if (_context?.GetStyle(styleId) is { } style)
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
        if (_tableExtractor is null)
            throw new InvalidOperationException("Parser is not initialized.");

        return _tableExtractor.ProcessTable(table);
    }

    /// <summary>
    /// Processes structured document tags (content controls)
    /// </summary>
    private DocumentNode? ProcessStructuredDocumentTag(SdtBlock sdtBlock)
    {
        var content = sdtBlock.SdtContentBlock;
        if (content is null) return null;

        // Extract content control properties
        var ccProperties = ExtractContentControlProperties(sdtBlock.SdtProperties);

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
                paraNode.Metadata[DocumentNode.IsSdtContentKey] = true;
                paraNode.ContentControlProperties = ccProperties;

                // Set the value in content control properties
                if (ccProperties is not null)
                {
                    ccProperties.Value = paraNode.GetText();
                }

                paraNode.AcceptAllChanges();
            }
            return paraNode;
        }

        // For SDT blocks with multiple elements, create a container node and store original XML
        var containerNode = new DocumentNode(ContentType.ContentControl, "")
        {
            OriginalXml = sdtBlock.OuterXml,
            Metadata = { [DocumentNode.IsSdtBlockKey] = true },
            ContentControlProperties = ccProperties
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
                // Mark the child as physically inside the SDT. The writer emits the SDT's own XML,
                // which already contains this content, so it must not write the child again — but it
                // must still write children that the heading hierarchy merely nested underneath,
                // which carry no such mark.
                childNode.Metadata[DocumentNode.IsInsideParentSdtKey] = true;
                containerNode.AddChild(childNode);
            }
        }

        // Set the value in content control properties from children
        if (ccProperties is not null && containerNode.Children.Count > 0)
        {
            ccProperties.Value = string.Join(" ", containerNode.Children.Select(c => c.GetText()).Where(t => !string.IsNullOrWhiteSpace(t)));
        }

        // Set Text for the container
        containerNode.Text = ccProperties?.Value ?? "";

        containerNode.AcceptAllChanges();

        // Always return the container for round-trip fidelity - even empty SDT blocks should be preserved
        return containerNode;
    }

    /// <summary>
    /// Processes inline structured document tags (content controls within a paragraph)
    /// </summary>
    private List<FormattedRun> ProcessInlineSdt(SdtRun sdtRun)
    {
        var result = new List<FormattedRun>();
        var ccProperties = ExtractContentControlProperties(sdtRun.SdtProperties);
        var content = sdtRun.SdtContentRun;

        if (content is null) return result;

        // Process runs within the SDT
        foreach (var run in content.Elements<Run>())
        {
            var runs = ProcessRun(run);
            // Mark all runs as belonging to a content control
            foreach (var r in runs)
            {
                // Always set content control properties on the run
                r.ContentControlProperties = ccProperties;

                // If this is a document property content control, also create a document property field
                if (ccProperties?.Type == ContentControlType.DocumentProperty &&
                    !string.IsNullOrEmpty(ccProperties.DataBindingXPath))
                {
                    var propName = DocumentPropertyHelpers.ExtractPropertyNameFromXPath(ccProperties.DataBindingXPath);
                    r.DocumentPropertyField = new DocumentPropertyField
                    {
                        PropertyName = propName,
                        PropertyType = DocumentPropertyHelpers.DeterminePropertyType(propName),
                        Value = r.Text,
                        FieldCode = ccProperties.DataBindingXPath
                    };
                }
            }
            result.AddRange(runs);
        }

        // Set the value in content control properties
        if (ccProperties is not null)
        {
            ccProperties.Value = string.Concat(result.Select(r => r.Text));
        }

        return result;
    }

    /// <summary>
    /// Extracts content control properties from an SdtProperties element
    /// </summary>
    private ContentControlProperties? ExtractContentControlProperties(SdtProperties? sdtPr)
    {
        if (sdtPr is null) return null;

        var props = new ContentControlProperties();

        // ID
        props.Id = sdtPr.GetFirstChild<SdtId>()?.Val?.Value;

        // Tag
        props.Tag = sdtPr.GetFirstChild<Tag>()?.Val?.Value;

        // Alias (Title)
        props.Alias = sdtPr.GetFirstChild<SdtAlias>()?.Val?.Value;

        // Lock settings - use fully qualified name to avoid ambiguity with System.Threading.Lock
        var lockSetting = sdtPr.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
        if (lockSetting is not null)
        {
            var lockVal = lockSetting.Val?.Value;
            props.LockContentControl = lockVal == LockingValues.SdtLocked || lockVal == LockingValues.SdtContentLocked;
            props.LockContents = lockVal == LockingValues.ContentLocked || lockVal == LockingValues.SdtContentLocked;
        }

        // Placeholder text - get from DocPartGallery or the placeholder's child element
        var placeholder = sdtPr.GetFirstChild<SdtPlaceholder>();
        if (placeholder is not null)
        {
            // The placeholder contains a DocPartGallery element
            var docPartGallery = placeholder.GetFirstChild<DocPartGallery>();
            if (docPartGallery?.Val?.Value is not null)
            {
                props.PlaceholderText = docPartGallery.Val.Value;
            }
        }

        // Show placeholder status
        props.ShowingPlaceholder = sdtPr.GetFirstChild<ShowingPlaceholder>() is not null;

        // Data binding
        var dataBinding = sdtPr.GetFirstChild<DataBinding>();
        if (dataBinding is not null)
        {
            props.DataBindingPrefixMappings = dataBinding.PrefixMappings?.Value;
            props.DataBindingXPath = dataBinding.XPath?.Value;
            props.DataBindingStoreItemId = dataBinding.StoreItemId?.Value;
        }

        // Appearance and Color are Word 2013+ features (w15 namespace)
        // Try to get them from extended attributes if present
        var w15Appearance = sdtPr.Descendants().FirstOrDefault(e => e.LocalName == "appearance");
        if (w15Appearance is not null)
        {
            var valAttr = w15Appearance.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
            if (!string.IsNullOrEmpty(valAttr.Value))
            {
                props.Appearance = valAttr.Value;
            }
        }

        var w15Color = sdtPr.Descendants().FirstOrDefault(e => e.LocalName == "color");
        if (w15Color is not null)
        {
            var valAttr = w15Color.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
            if (!string.IsNullOrEmpty(valAttr.Value))
            {
                props.Color = valAttr.Value;
            }
        }

        // Determine content control type based on specific elements
        props.Type = DetermineContentControlType(sdtPr);

        // Type-specific properties
        switch (props.Type)
        {
            case ContentControlType.Date:
                var datePr = sdtPr.GetFirstChild<SdtContentDate>();
                if (datePr is not null)
                {
                    props.DateFormat = datePr.DateFormat?.Val?.Value;
                    // Get locale from language element
                    var langElem = datePr.Descendants().FirstOrDefault(e => e.LocalName == "lid");
                    if (langElem is not null)
                    {
                        var valAttr = langElem.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
                        props.DateLocale = valAttr.Value;
                    }
                    // Get the date value
                    if (datePr.FullDate is not null)
                    {
                        props.DateValue = datePr.FullDate.Value;
                    }
                }
                break;

            case ContentControlType.DropDownList:
                var dropDown = sdtPr.GetFirstChild<SdtContentDropDownList>();
                if (dropDown is not null)
                {
                    foreach (var item in dropDown.Elements<ListItem>())
                    {
                        props.ListItems.Add(new ContentControlListItem
                        {
                            DisplayText = item.DisplayText?.Value,
                            Value = item.Value?.Value
                        });
                    }
                }
                break;

            case ContentControlType.ComboBox:
                var comboBox = sdtPr.GetFirstChild<SdtContentComboBox>();
                if (comboBox is not null)
                {
                    foreach (var item in comboBox.Elements<ListItem>())
                    {
                        props.ListItems.Add(new ContentControlListItem
                        {
                            DisplayText = item.DisplayText?.Value,
                            Value = item.Value?.Value
                        });
                    }
                }
                break;

            case ContentControlType.Checkbox:
                // Checkbox state is in w14:checkbox element
                var checkbox = sdtPr.Descendants().FirstOrDefault(e => e.LocalName == "checkbox");
                if (checkbox is not null)
                {
                    var checkedState = checkbox.Descendants().FirstOrDefault(e => e.LocalName == "checked");
                    if (checkedState is not null)
                    {
                        var valAttr = checkedState.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
                        props.IsChecked = valAttr.Value == "1" || valAttr.Value?.ToLower() == "true";
                    }
                }
                break;
        }

        return props;
    }

    /// <summary>
    /// Determines the type of content control from its properties
    /// </summary>
    private static ContentControlType DetermineContentControlType(SdtProperties sdtPr)
    {
        // Check for specific content control type elements
        if (sdtPr.GetFirstChild<SdtContentRichText>() is not null)
            return ContentControlType.RichText;
        if (sdtPr.GetFirstChild<SdtContentText>() is not null)
            return ContentControlType.PlainText;
        if (sdtPr.GetFirstChild<SdtContentPicture>() is not null)
            return ContentControlType.Picture;
        if (sdtPr.GetFirstChild<SdtContentDate>() is not null)
            return ContentControlType.Date;
        if (sdtPr.GetFirstChild<SdtContentDropDownList>() is not null)
            return ContentControlType.DropDownList;
        if (sdtPr.GetFirstChild<SdtContentComboBox>() is not null)
            return ContentControlType.ComboBox;
        if (sdtPr.GetFirstChild<SdtContentGroup>() is not null)
            return ContentControlType.Group;
        if (sdtPr.GetFirstChild<SdtContentBibliography>() is not null)
            return ContentControlType.Bibliography;
        if (sdtPr.GetFirstChild<SdtContentCitation>() is not null)
            return ContentControlType.Citation;
        if (sdtPr.GetFirstChild<SdtContentEquation>() is not null)
            return ContentControlType.Equation;

        // Check for checkbox (w14:checkbox)
        if (sdtPr.Descendants().Any(e => e.LocalName == "checkbox"))
            return ContentControlType.Checkbox;

        // Check for document property binding
        var dataBinding = sdtPr.GetFirstChild<DataBinding>();
        if (dataBinding?.XPath?.Value is not null)
        {
            var xpath = dataBinding.XPath.Value;
            if (xpath.Contains("coreProperties") || xpath.Contains("extended-properties") ||
                xpath.Contains("custom-properties"))
            {
                return ContentControlType.DocumentProperty;
            }
        }

        // Check for repeating section
        if (sdtPr.Descendants().Any(e => e.LocalName == "repeatingSection"))
            return ContentControlType.RepeatingSection;
        if (sdtPr.Descendants().Any(e => e.LocalName == "repeatingSectionItem"))
            return ContentControlType.RepeatingSectionItem;

        return ContentControlType.Unknown;
    }

    /// <summary>Releases resources used by the parser.</summary>
    public void Dispose()
    {
        ReleaseParsingState();
        _packageBytes = null;
        GC.SuppressFinalize(this);
    }
}
