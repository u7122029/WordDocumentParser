using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using WordDocumentParser.Models.Images;
using WordDocumentParser.Models.Package;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using XP = DocumentFormat.OpenXml.ExtendedProperties;

namespace WordDocumentParser.Writing;

/// <summary>
/// Builds one <c>.docx</c> package from a document tree. All state is local to the session, so
/// nothing carries over between writes.
/// </summary>
/// <remarks>
/// <para>
/// When the document carries the bytes of the package it was parsed from, the session edits a copy
/// of that package: it replaces the body with the current tree and overwrites only those parts the
/// caller changed. Everything else — comments, revisions, embedded objects, the numbering the body
/// references, the hyperlink relationships a header owns, and every extension-namespace element —
/// travels through untouched, because it is never taken apart in the first place.
/// </para>
/// <para>
/// Reassembling a package part by part could only preserve what the model happened to capture, and
/// silently dropped the rest. That path now runs only for documents built in code, which have no
/// source package to preserve.
/// </para>
/// </remarks>
internal sealed class DocumentWriteSession(WordDocument document, RecoveryOptions recovery)
{
    private readonly Dictionary<string, string> _relationshipMapping = [];
    private readonly Dictionary<string, string> _headerRelationshipMapping = [];
    private readonly Dictionary<string, string> _footerRelationshipMapping = [];

    private DocumentPackageData _packageData = document.PackageData;
    private WordprocessingDocument? _package;
    private MainDocumentPart? _mainPart;
    private Body? _body;
    private NumberingDefinitionsPart? _numberingPart;
    private TableEditor _tableEditor = null!;
    private int _currentListId = 1;
    private uint _imageCounter = 1;

    /// <summary>
    /// Produces the complete package bytes.
    /// </summary>
    /// <returns>The <c>.docx</c> package.</returns>
    /// <remarks>
    /// The package is assembled fully in memory before any destination is touched, so a failure part
    /// way through cannot leave a half-written file behind.
    /// </remarks>
    public byte[] Build()
    {
        document.SyncCustomPropertiesToXml();
        _packageData = document.PackageData;
        _tableEditor = new TableEditor(UpdateRelationshipIds);

        using var buffer = new MemoryStream();

        if (_packageData.CanEditInPlace)
        {
            EditExistingPackage(buffer);
        }
        else
        {
            CreateNewPackage(buffer);
        }

        return buffer.ToArray();
    }

    #region Edit an existing package

    private void EditExistingPackage(MemoryStream buffer)
    {
        buffer.Write(_packageData.OriginalPackageBytes!, 0, _packageData.OriginalPackageBytes!.Length);
        buffer.Position = 0;

        using var package = WordprocessingDocument.Open(buffer, true);
        _package = package;
        _mainPart = package.MainDocumentPart
                    ?? throw new DocumentPreservationException(
                        "word/document.xml", "The source package has no main document part.");

        var body = _mainPart.Document?.Body
                   ?? throw new DocumentPreservationException(
                       "word/document.xml", "The source package has no document body.");

        _numberingPart = _mainPart.NumberingDefinitionsPart;

        AddResourcesMissingFromPackage();
        ReplaceBody(body);
        OverwriteChangedParts();
        ApplyDocumentProperties();

        package.Save();
    }

    /// <summary>
    /// Rebuilds the body from the tree, keeping the section properties and any body-level elements
    /// the tree does not model, such as bookmarks that span paragraphs.
    /// </summary>
    private void ReplaceBody(Body body)
    {
        var sectionProperties = body.Elements<SectionProperties>().LastOrDefault();
        sectionProperties?.Remove();

        // Body children the tree has no node for. Keeping them, at roughly their original position,
        // is what stops a save from dropping bookmarks and revision markers anchored between blocks.
        var unmodelled = new List<(int Index, OpenXmlElement Element)>();
        var index = 0;
        foreach (var child in body.ChildElements)
        {
            if (child is not (Paragraph or Table or SdtBlock))
            {
                unmodelled.Add((index, child));
            }
            index++;
        }

        body.RemoveAllChildren();
        _body = body;

        ProcessNode(document.Root);

        foreach (var (originalIndex, element) in unmodelled)
        {
            if (originalIndex < body.ChildElements.Count)
            {
                body.InsertAt(element, originalIndex);
            }
            else
            {
                body.Append(element);
            }
        }

        if (sectionProperties is not null)
        {
            body.Append(sectionProperties);
        }
        else
        {
            AddDefaultSectionProperties();
        }
    }

    /// <summary>
    /// Adds media and hyperlink relationships the model gained since the package was parsed — the
    /// result of merging another document in — and maps their placeholder IDs onto the real ones.
    /// </summary>
    private void AddResourcesMissingFromPackage()
    {
        var existing = _mainPart!.Parts
            .Select(pair => pair.RelationshipId)
            .Concat(_mainPart.HyperlinkRelationships.Select(rel => rel.Id))
            .ToHashSet(StringComparer.Ordinal);

        foreach (var (relationshipId, imageData) in _packageData.Images)
        {
            if (existing.Contains(relationshipId)) continue;

            var imagePart = _mainPart.AddImagePart(NormalizeImageContentType(imageData.ContentType));
            using (var stream = new MemoryStream(imageData.Data, writable: false))
            {
                imagePart.FeedData(stream);
            }

            _relationshipMapping[relationshipId] = _mainPart.GetIdOfPart(imagePart);
        }

        foreach (var (relationshipId, hyperlink) in _packageData.HyperlinkRelationships)
        {
            if (existing.Contains(relationshipId)) continue;

            if (!Uri.TryCreate(hyperlink.Url, UriKind.RelativeOrAbsolute, out var uri))
            {
                recovery.Report(relationshipId, $"Hyperlink target '{hyperlink.Url}' is not a valid URI.");
                continue;
            }

            var relationship = _mainPart.AddHyperlinkRelationship(uri, hyperlink.IsExternal);
            _relationshipMapping[relationshipId] = relationship.Id;
        }
    }

    /// <summary>
    /// Overwrites the package's XML parts, but only where the current value differs from what was
    /// parsed.
    /// </summary>
    /// <remarks>
    /// Parts are written as raw XML rather than through the SDK's typed constructors. The typed path
    /// forced the previous implementation to strip the <c>w14</c>, <c>w15</c>, and <c>w16</c>
    /// namespaces wholesale to get past validation, which deleted real content: a checkbox lost the
    /// <c>w14:checkbox</c> element that holds its state.
    /// </remarks>
    private void OverwriteChangedParts()
    {
        var baseline = _packageData.Baseline;

        WritePartIfChanged(_packageData.StylesXml, baseline?.StylesXml,
            () => _mainPart!.StyleDefinitionsPart ?? _mainPart!.AddNewPart<StyleDefinitionsPart>());

        WritePartIfChanged(_packageData.ThemeXml, baseline?.ThemeXml,
            () => _mainPart!.ThemePart ?? _mainPart!.AddNewPart<ThemePart>());

        WritePartIfChanged(_packageData.FontTableXml, baseline?.FontTableXml,
            () => _mainPart!.FontTablePart ?? _mainPart!.AddNewPart<FontTablePart>());

        WritePartIfChanged(_packageData.NumberingXml, baseline?.NumberingXml,
            () => _numberingPart ??= _mainPart!.NumberingDefinitionsPart ?? _mainPart!.AddNewPart<NumberingDefinitionsPart>());

        WritePartIfChanged(_packageData.SettingsXml, baseline?.SettingsXml,
            () => _mainPart!.DocumentSettingsPart ?? _mainPart!.AddNewPart<DocumentSettingsPart>());

        WritePartIfChanged(_packageData.WebSettingsXml, baseline?.WebSettingsXml,
            () => _mainPart!.WebSettingsPart ?? _mainPart!.AddNewPart<WebSettingsPart>());

        WritePartIfChanged(_packageData.FootnotesXml, baseline?.FootnotesXml,
            () => _mainPart!.FootnotesPart ?? _mainPart!.AddNewPart<FootnotesPart>());

        WritePartIfChanged(_packageData.EndnotesXml, baseline?.EndnotesXml,
            () => _mainPart!.EndnotesPart ?? _mainPart!.AddNewPart<EndnotesPart>());

        WritePartIfChanged(_packageData.GlossaryDocumentXml, baseline?.GlossaryDocumentXml,
            () => _mainPart!.GlossaryDocumentPart ?? _mainPart!.AddNewPart<GlossaryDocumentPart>());

        foreach (var (relationshipId, xml) in _packageData.Headers)
        {
            if (baseline is not null && baseline.Headers.TryGetValue(relationshipId, out var original) && original == xml)
                continue;

            if (_mainPart!.GetPartById(relationshipId) is HeaderPart headerPart)
            {
                WriteXmlToPart(headerPart, xml);
            }
        }

        foreach (var (relationshipId, xml) in _packageData.Footers)
        {
            if (baseline is not null && baseline.Footers.TryGetValue(relationshipId, out var original) && original == xml)
                continue;

            if (_mainPart!.GetPartById(relationshipId) is FooterPart footerPart)
            {
                WriteXmlToPart(footerPart, xml);
            }
        }
    }

    private void WritePartIfChanged(string? current, string? baseline, Func<OpenXmlPart> getPart)
    {
        if (string.IsNullOrEmpty(current) || current == baseline) return;

        WriteXmlToPart(getPart(), current);
    }

    #endregion

    #region Create a new package

    private void CreateNewPackage(MemoryStream buffer)
    {
        using var package = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document, autoSave: false);
        _package = package;
        _mainPart = package.AddMainDocumentPart();
        _mainPart.Document = new Document();
        _body = new Body();
        _mainPart.Document.Append(_body);
        _numberingPart = null;

        RestoreModelledParts();
        ProcessNode(document.Root);
        AddSectionPropertiesFromModel();
        ApplyDocumentProperties();

        package.Save();
    }

    /// <summary>
    /// Recreates the parts the model captured, for documents that have package data but no source
    /// package to copy. XML is written verbatim rather than through typed constructors, so extension
    /// namespaces survive.
    /// </summary>
    private void RestoreModelledParts()
    {
        WriteRawPart(_packageData.StylesXml, () => _mainPart!.AddNewPart<StyleDefinitionsPart>());
        WriteRawPart(_packageData.ThemeXml, () => _mainPart!.AddNewPart<ThemePart>());
        WriteRawPart(_packageData.FontTableXml, () => _mainPart!.AddNewPart<FontTablePart>());
        WriteRawPart(_packageData.SettingsXml, () => _mainPart!.AddNewPart<DocumentSettingsPart>());
        WriteRawPart(_packageData.WebSettingsXml, () => _mainPart!.AddNewPart<WebSettingsPart>());
        WriteRawPart(_packageData.FootnotesXml, () => _mainPart!.AddNewPart<FootnotesPart>());
        WriteRawPart(_packageData.EndnotesXml, () => _mainPart!.AddNewPart<EndnotesPart>());

        if (!string.IsNullOrEmpty(_packageData.NumberingXml))
        {
            _numberingPart = _mainPart!.AddNewPart<NumberingDefinitionsPart>();
            WriteXmlToPart(_numberingPart, _packageData.NumberingXml);
        }

        if (_packageData.StylesXml is null)
        {
            AddDefaultStyles();
        }

        RestoreImages();
        RestoreHyperlinkRelationships();
        RestoreHeadersAndFooters();
        RestoreCustomXmlParts();
        RestoreGlossary();
    }

    private void WriteRawPart(string? xml, Func<OpenXmlPart> createPart)
    {
        if (string.IsNullOrEmpty(xml)) return;
        WriteXmlToPart(createPart(), xml);
    }

    private void RestoreImages()
    {
        var mainPart = _mainPart!;

        foreach (var (relationshipId, imageData) in _packageData.Images)
        {
            var imagePart = mainPart.AddImagePart(NormalizeImageContentType(imageData.ContentType));
            using (var stream = new MemoryStream(imageData.Data, writable: false))
            {
                imagePart.FeedData(stream);
            }

            _relationshipMapping[relationshipId] = mainPart.GetIdOfPart(imagePart);
        }
    }

    private void RestoreHyperlinkRelationships()
    {
        foreach (var (relationshipId, hyperlink) in _packageData.HyperlinkRelationships)
        {
            if (!Uri.TryCreate(hyperlink.Url, UriKind.RelativeOrAbsolute, out var uri))
            {
                recovery.Report(relationshipId, $"Hyperlink target '{hyperlink.Url}' is not a valid URI.");
                continue;
            }

            var relationship = _mainPart!.AddHyperlinkRelationship(uri, hyperlink.IsExternal);
            _relationshipMapping[relationshipId] = relationship.Id;
        }
    }

    private void RestoreHeadersAndFooters()
    {
        foreach (var (originalId, xml) in _packageData.Headers)
        {
            var headerPart = _mainPart!.AddNewPart<HeaderPart>();
            var content = xml;

            if (_packageData.HeaderImages.TryGetValue(originalId, out var images))
            {
                content = RestorePartImages(headerPart, images, content);
            }

            WriteXmlToPart(headerPart, content);
            _headerRelationshipMapping[originalId] = _mainPart.GetIdOfPart(headerPart);
        }

        foreach (var (originalId, xml) in _packageData.Footers)
        {
            var footerPart = _mainPart!.AddNewPart<FooterPart>();
            var content = xml;

            if (_packageData.FooterImages.TryGetValue(originalId, out var images))
            {
                content = RestorePartImages(footerPart, images, content);
            }

            WriteXmlToPart(footerPart, content);
            _footerRelationshipMapping[originalId] = _mainPart.GetIdOfPart(footerPart);
        }
    }

    private static string RestorePartImages<TPart>(
        TPart part, Dictionary<string, ImagePartData> images, string xml)
        where TPart : OpenXmlPartContainer, ISupportedRelationship<ImagePart>
    {
        var mapping = new Dictionary<string, string>();

        foreach (var (originalId, imageData) in images)
        {
            var imagePart = part.AddImagePart(NormalizeImageContentType(imageData.ContentType));
            using (var stream = new MemoryStream(imageData.Data, writable: false))
            {
                imagePart.FeedData(stream);
            }
            mapping[originalId] = part.GetIdOfPart(imagePart);
        }

        return ApplyRelationshipMapping(xml, mapping);
    }

    private void RestoreCustomXmlParts()
    {
        foreach (var (uri, data) in _packageData.CustomXmlParts)
        {
            try
            {
                var customXmlPart = _mainPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                WriteXmlToPart(customXmlPart, data.XmlContent);

                if (!string.IsNullOrEmpty(data.PropertiesXml))
                {
                    WriteXmlToPart(customXmlPart.AddNewPart<CustomXmlPropertiesPart>(), data.PropertiesXml);
                }
            }
            catch (Exception ex)
            {
                recovery.Report(uri, "Custom XML part could not be restored.", ex);
            }
        }
    }

    private void RestoreGlossary()
    {
        if (string.IsNullOrEmpty(_packageData.GlossaryDocumentXml)) return;

        try
        {
            var glossaryPart = _mainPart!.AddNewPart<GlossaryDocumentPart>();
            var xml = _packageData.GlossaryDocumentXml;

            if (_packageData.GlossaryImages.Count > 0)
            {
                xml = RestorePartImages(glossaryPart, _packageData.GlossaryImages, xml);
            }

            WriteXmlToPart(glossaryPart, xml);

            if (!string.IsNullOrEmpty(_packageData.GlossaryStylesXml))
            {
                WriteXmlToPart(glossaryPart.AddNewPart<StyleDefinitionsPart>(), _packageData.GlossaryStylesXml);
            }

            if (!string.IsNullOrEmpty(_packageData.GlossaryFontTableXml))
            {
                WriteXmlToPart(glossaryPart.AddNewPart<FontTablePart>(), _packageData.GlossaryFontTableXml);
            }
        }
        catch (Exception ex)
        {
            recovery.Report("word/glossary/document.xml", "Glossary document could not be restored.", ex);
        }
    }

    private void AddSectionPropertiesFromModel()
    {
        if (_packageData.SectionPropertiesXml.Count > 0)
        {
            var sectionProperties = new SectionProperties(_packageData.SectionPropertiesXml[^1]);
            RemapHeaderFooterReferences(sectionProperties);
            _body!.Append(sectionProperties);
            return;
        }

        AddDefaultSectionProperties();
    }

    private void AddDefaultSectionProperties()
    {
        var sectionProperties = new SectionProperties();
        sectionProperties.Append(new PageSize { Width = 12240, Height = 15840 });
        sectionProperties.Append(new PageMargin
        {
            Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0
        });
        _body!.Append(sectionProperties);
    }

    private void RemapHeaderFooterReferences(SectionProperties sectionProperties)
    {
        foreach (var reference in sectionProperties.Elements<HeaderReference>())
        {
            if (reference.Id?.Value is { } id && _headerRelationshipMapping.TryGetValue(id, out var newId))
            {
                reference.Id = newId;
            }
        }

        foreach (var reference in sectionProperties.Elements<FooterReference>())
        {
            if (reference.Id?.Value is { } id && _footerRelationshipMapping.TryGetValue(id, out var newId))
            {
                reference.Id = newId;
            }
        }
    }

    #endregion

    #region Tree traversal

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
            case ContentType.Paragraph:
            case ContentType.ListItem:
            case ContentType.HyperlinkText:
            case ContentType.TextRun:
            case ContentType.ContentControl:
                WriteBlockNode(node);
                break;

            case ContentType.Table:
                WriteTable(node);
                WriteChildren(node);
                break;

            case ContentType.Image:
                WriteImage(node);
                break;

            case ContentType.List:
                WriteChildren(node);
                break;
        }
    }

    /// <summary>
    /// Writes a paragraph-like node, then the content nested beneath it.
    /// </summary>
    private void WriteBlockNode(DocumentNode node)
    {
        if (node.Type == ContentType.ListItem)
        {
            EnsureNumberingPart();
        }

        var element = BuildBlockElement(node);
        if (element is not null)
        {
            _body!.Append(element);
        }

        WriteChildren(node);
    }

    /// <summary>
    /// Writes a node's children, skipping those whose content is already inside the node's own SDT
    /// XML.
    /// </summary>
    /// <remarks>
    /// The distinction matters: a heading wrapped in a content control still gathers the rest of its
    /// section as children by the heading hierarchy. Skipping every child of an SDT node would
    /// discard that section; skipping none would duplicate the control's own content.
    /// </remarks>
    private void WriteChildren(DocumentNode node)
    {
        var emittedWithParent = !string.IsNullOrEmpty(node.OriginalXml);

        foreach (var child in node.Children)
        {
            if (child.IsInsideParentSdt) continue;

            // An image child parsed out of this paragraph is already present in the paragraph's XML.
            if (emittedWithParent && child.Type == ContentType.Image) continue;

            ProcessNode(child);
        }
    }

    /// <summary>
    /// Produces the XML element for a paragraph-like node: its original XML with the caller's edits
    /// applied, or a freshly built paragraph when the node has no original.
    /// </summary>
    private OpenXmlElement? BuildBlockElement(DocumentNode node)
    {
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            var xml = UpdateRelationshipIds(node.OriginalXml);

            OpenXmlElement element = xml.TrimStart().StartsWith("<w:sdt", StringComparison.Ordinal)
                ? new SdtBlock(xml)
                : new Paragraph(xml);

            ParagraphEditor.Apply(element, node);
            return element;
        }

        // A content control container holds its blocks as children rather than as content of its
        // own. Once it has no XML — because the control was removed — there is nothing to emit for
        // the container itself, and its children are written in its place.
        if (node.Type == ContentType.ContentControl)
        {
            return null;
        }

        return BuildParagraphFromModel(node);
    }

    private Paragraph BuildParagraphFromModel(DocumentNode node)
    {
        var paragraph = new Paragraph();
        var formatting = node.ParagraphFormatting;

        if (formatting is not null)
        {
            // A node built in code has every assigned property marked as changed, so this writes
            // exactly what the caller set.
            ParagraphEditor.ApplyParagraphFormatting(paragraph, formatting);
        }

        EnsureStructuralParagraphProperties(paragraph, node);

        if (node.HasFormattedRuns)
        {
            foreach (var modelRun in node.Runs)
            {
                paragraph.Append(BuildRun(modelRun));
            }
        }
        else if (!string.IsNullOrEmpty(node.Text))
        {
            paragraph.Append(new Run(new Text(node.Text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        return paragraph;
    }

    /// <summary>
    /// Adds the style and numbering a heading or list item needs when the caller did not set them.
    /// </summary>
    private void EnsureStructuralParagraphProperties(Paragraph paragraph, DocumentNode node)
    {
        if (node.Type is not (ContentType.Heading or ContentType.ListItem)) return;

        var props = paragraph.ParagraphProperties;
        if (props is null)
        {
            props = new ParagraphProperties();
            paragraph.InsertAt(props, 0);
        }

        if (props.ParagraphStyleId is null)
        {
            OoxmlOrder.InsertParagraphProperty(props, new ParagraphStyleId
            {
                Val = node.Type == ContentType.Heading ? $"Heading{node.HeadingLevel}" : "ListParagraph"
            });
        }

        if (node.Type == ContentType.ListItem && props.NumberingProperties is null)
        {
            var level = node.Metadata.TryGetValue("ListLevel", out var value) ? Convert.ToInt32(value) : 0;
            OoxmlOrder.InsertParagraphProperty(props, new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = _currentListId }));
        }
    }

    private static Run BuildRun(Models.Formatting.FormattedRun modelRun)
    {
        var run = new Run();

        if (modelRun.Formatting.HasChanges || modelRun.Formatting.HasFormatting)
        {
            RunEditor.ApplyRunFormatting(run, MarkAllForWrite(modelRun.Formatting));
        }

        if (modelRun.IsTab)
        {
            run.Append(new TabChar());
        }
        else if (modelRun.IsBreak)
        {
            var lineBreak = new Break();
            if (OoxmlEnum.Parse<BreakValues>(modelRun.BreakType) is { } breakType)
            {
                lineBreak.Type = breakType;
            }
            run.Append(lineBreak);
        }
        else if (!string.IsNullOrEmpty(modelRun.Text))
        {
            run.Append(new Text(modelRun.Text) { Space = SpaceProcessingModeValues.Preserve });
        }

        return run;
    }

    /// <summary>
    /// Marks every populated property of a run's formatting as changed, so a run being written from
    /// scratch emits all of it rather than only what was assigned since the last baseline.
    /// </summary>
    private static Models.Formatting.RunFormatting MarkAllForWrite(Models.Formatting.RunFormatting formatting)
    {
        var copy = formatting.Clone();

        foreach (var name in RunFormattingPropertyNames)
        {
            copy.MarkChanged(name);
        }

        return copy;
    }

    private static readonly string[] RunFormattingPropertyNames =
    [
        nameof(Models.Formatting.RunFormatting.Bold), nameof(Models.Formatting.RunFormatting.Italic),
        nameof(Models.Formatting.RunFormatting.Underline), nameof(Models.Formatting.RunFormatting.UnderlineStyle),
        nameof(Models.Formatting.RunFormatting.Strike), nameof(Models.Formatting.RunFormatting.DoubleStrike),
        nameof(Models.Formatting.RunFormatting.FontFamily), nameof(Models.Formatting.RunFormatting.FontFamilyAscii),
        nameof(Models.Formatting.RunFormatting.FontFamilyEastAsia), nameof(Models.Formatting.RunFormatting.FontFamilyComplexScript),
        nameof(Models.Formatting.RunFormatting.FontSize), nameof(Models.Formatting.RunFormatting.FontSizeComplexScript),
        nameof(Models.Formatting.RunFormatting.Color), nameof(Models.Formatting.RunFormatting.Highlight),
        nameof(Models.Formatting.RunFormatting.Superscript), nameof(Models.Formatting.RunFormatting.Subscript),
        nameof(Models.Formatting.RunFormatting.SmallCaps), nameof(Models.Formatting.RunFormatting.AllCaps),
        nameof(Models.Formatting.RunFormatting.Shading), nameof(Models.Formatting.RunFormatting.StyleId)
    ];

    private void WriteTable(DocumentNode node)
    {
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            var table = new Table(UpdateRelationshipIds(node.OriginalXml));

            if (node.GetTableData() is { } data)
            {
                _tableEditor.Apply(table, data);
            }

            _body!.Append(table);
            return;
        }

        var tableData = node.GetTableData();
        if (tableData is null || tableData.Rows.Count == 0)
        {
            _body!.Append(new Paragraph(new Run(new Text("[Table]"))));
            return;
        }

        _body!.Append(TableBuilder.Build(tableData));
    }

    private void WriteImage(DocumentNode node)
    {
        var imageData = node.GetImageData();
        if (imageData?.Data is null || imageData.Data.Length == 0)
        {
            _body!.Append(new Paragraph(new Run(new Text(node.Text.Length > 0 ? node.Text : "[Image]"))));
            return;
        }

        string relationshipId;
        if (!string.IsNullOrEmpty(imageData.Id) && _relationshipMapping.TryGetValue(imageData.Id, out var mapped))
        {
            relationshipId = mapped;
        }
        else if (!string.IsNullOrEmpty(imageData.Id) && PartExists(imageData.Id))
        {
            relationshipId = imageData.Id;
        }
        else
        {
            var mainPart = _mainPart!;
            var imagePart = mainPart.AddImagePart(NormalizeImageContentType(imageData.ContentType));
            using (var stream = new MemoryStream(imageData.Data, writable: false))
            {
                imagePart.FeedData(stream);
            }
            relationshipId = mainPart.GetIdOfPart(imagePart);
        }

        var widthEmu = imageData.WidthEmu > 0 ? imageData.WidthEmu : (long)(imageData.WidthInches * 914400);
        var heightEmu = imageData.HeightEmu > 0 ? imageData.HeightEmu : (long)(imageData.HeightInches * 914400);
        if (widthEmu <= 0) widthEmu = 914400 * 4;
        if (heightEmu <= 0) heightEmu = 914400 * 3;

        var drawing = CreateImageDrawing(relationshipId, widthEmu, heightEmu, imageData);
        _body!.Append(new Paragraph(new Run(drawing)));
    }

    private bool PartExists(string relationshipId)
    {
        try
        {
            return _mainPart!.GetPartById(relationshipId) is not null;
        }
        catch (ArgumentOutOfRangeException)
        {
            return false;
        }
    }

    private Drawing CreateImageDrawing(string relationshipId, long widthEmu, long heightEmu, ImageData imageData)
    {
        var imageId = _imageCounter++;
        var formatting = imageData.Formatting;

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
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

        if (formatting is null || formatting.IsInline)
        {
            return new Drawing(new DW.Inline(
                extent,
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                docProperties,
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                graphic)
            {
                DistanceFromTop = (uint)(formatting?.DistanceFromTop ?? 0),
                DistanceFromBottom = (uint)(formatting?.DistanceFromBottom ?? 0),
                DistanceFromLeft = (uint)(formatting?.DistanceFromLeft ?? 0),
                DistanceFromRight = (uint)(formatting?.DistanceFromRight ?? 0)
            });
        }

        var anchor = new DW.Anchor
        {
            DistanceFromTop = (uint)(formatting.DistanceFromTop ?? 0),
            DistanceFromBottom = (uint)(formatting.DistanceFromBottom ?? 0),
            DistanceFromLeft = (uint)(formatting.DistanceFromLeft ?? 0),
            DistanceFromRight = (uint)(formatting.DistanceFromRight ?? 0),
            SimplePos = false,
            RelativeHeight = (uint)(formatting.RelativeHeight ?? 0),
            BehindDoc = formatting.BehindDocument,
            Locked = formatting.Locked,
            LayoutInCell = formatting.LayoutInCell,
            AllowOverlap = formatting.AllowOverlap
        };

        anchor.Append(new DW.SimplePosition { X = 0, Y = 0 });

        var horizontalPosition = new DW.HorizontalPosition
        {
            RelativeFrom = formatting.HorizontalRelativeTo switch
            {
                "Page" => DW.HorizontalRelativePositionValues.Page,
                "Margin" => DW.HorizontalRelativePositionValues.Margin,
                _ => DW.HorizontalRelativePositionValues.Column
            }
        };
        horizontalPosition.Append(new DW.PositionOffset(
            long.TryParse(formatting.HorizontalPosition, out var hOffset) ? hOffset.ToString() : "0"));
        anchor.Append(horizontalPosition);

        var verticalPosition = new DW.VerticalPosition
        {
            RelativeFrom = formatting.VerticalRelativeTo switch
            {
                "Page" => DW.VerticalRelativePositionValues.Page,
                "Margin" => DW.VerticalRelativePositionValues.Margin,
                _ => DW.VerticalRelativePositionValues.Paragraph
            }
        };
        verticalPosition.Append(new DW.PositionOffset(
            long.TryParse(formatting.VerticalPosition, out var vOffset) ? vOffset.ToString() : "0"));
        anchor.Append(verticalPosition);

        anchor.Append(extent);
        anchor.Append(new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 });

        anchor.Append(formatting.WrapType switch
        {
            "Square" => new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
            "Tight" => new DW.WrapTight { WrapText = DW.WrapTextValues.BothSides },
            "TopAndBottom" => (OpenXmlElement)new DW.WrapTopBottom(),
            _ => new DW.WrapNone()
        });

        anchor.Append(docProperties);
        anchor.Append(new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }));
        anchor.Append(graphic);

        return new Drawing(anchor);
    }

    #endregion

    #region Shared helpers

    /// <summary>
    /// Rewrites relationship IDs in a node's XML to the IDs its resources were given in this package.
    /// </summary>
    private string UpdateRelationshipIds(string xml) => ApplyRelationshipMapping(xml, _relationshipMapping);

    private static string ApplyRelationshipMapping(string xml, Dictionary<string, string> mapping)
    {
        if (mapping.Count == 0) return xml;

        var result = xml;
        foreach (var (oldId, newId) in mapping)
        {
            result = result
                .Replace($"r:embed=\"{oldId}\"", $"r:embed=\"{newId}\"")
                .Replace($"r:link=\"{oldId}\"", $"r:link=\"{newId}\"")
                .Replace($"r:id=\"{oldId}\"", $"r:id=\"{newId}\"");
        }

        return result;
    }

    /// <summary>
    /// Writes XML to a part verbatim, bypassing the typed API so extension elements and namespace
    /// declarations survive exactly as they were.
    /// </summary>
    private static void WriteXmlToPart(OpenXmlPart part, string xml)
    {
        var encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, encoding);

        if (!xml.TrimStart().StartsWith("<?xml", StringComparison.OrdinalIgnoreCase))
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        }

        writer.Write(xml);
    }

    private static string NormalizeImageContentType(string contentType) => contentType.ToLowerInvariant() switch
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

    private void EnsureNumberingPart()
    {
        if (_numberingPart is not null) return;

        _numberingPart = _mainPart!.NumberingDefinitionsPart ?? _mainPart.AddNewPart<NumberingDefinitionsPart>();
        if (_numberingPart.Numbering is not null) return;

        _numberingPart.Numbering = DefaultParts.CreateBulletNumbering();
    }

    private void AddDefaultStyles()
    {
        var stylesPart = _mainPart!.StyleDefinitionsPart ?? _mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = DefaultParts.CreateStyles();
    }

    private void ApplyDocumentProperties()
    {
        ApplyCoreProperties();
        ApplyExtendedProperties();
        ApplyCustomProperties();
    }

    /// <summary>
    /// Writes the core properties, applying clears as well as values.
    /// </summary>
    /// <remarks>
    /// A property the caller assigned is always written, including when the new value is null: when
    /// an existing package is being edited, skipping nulls left the original value in place, so
    /// removing a property reported success and changed nothing. Properties the caller never touched
    /// are left alone so an untouched document round-trips unchanged.
    /// </remarks>
    private void ApplyCoreProperties()
    {
        if (_packageData.CoreProperties is not { } core) return;

        var properties = _package!.PackageProperties;
        var writeAll = !_packageData.CanEditInPlace;

        Apply(nameof(core.Title), core.Title, value => properties.Title = value);
        Apply(nameof(core.Subject), core.Subject, value => properties.Subject = value);
        Apply(nameof(core.Creator), core.Creator, value => properties.Creator = value);
        Apply(nameof(core.Keywords), core.Keywords, value => properties.Keywords = value);
        Apply(nameof(core.Description), core.Description, value => properties.Description = value);
        Apply(nameof(core.Category), core.Category, value => properties.Category = value);
        Apply(nameof(core.LastModifiedBy), core.LastModifiedBy, value => properties.LastModifiedBy = value);
        Apply(nameof(core.Revision), core.Revision, value => properties.Revision = value);
        Apply(nameof(core.ContentStatus), core.ContentStatus, value => properties.ContentStatus = value);

        Apply(nameof(core.Created), core.Created, value => properties.Created = ParseTimestamp(value));
        Apply(nameof(core.Modified), core.Modified, value => properties.Modified = ParseTimestamp(value));

        void Apply(string propertyName, string? value, Action<string?> assign)
        {
            if (core.IsChanged(propertyName) || (writeAll && value is not null))
            {
                assign(value);
            }
        }
    }

    private static DateTime? ParseTimestamp(string? value) =>
        !string.IsNullOrEmpty(value) &&
        DateTime.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.RoundtripKind, out var parsed)
            ? parsed
            : null;

    private void ApplyExtendedProperties()
    {
        if (_packageData.ExtendedProperties is not { } extended) return;

        var part = _package!.ExtendedFilePropertiesPart;
        if (part is null)
        {
            part = _package.AddExtendedFilePropertiesPart();
            part.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
        }

        part.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
        var properties = part.Properties;
        var writeAll = !_packageData.CanEditInPlace;
        var culture = System.Globalization.CultureInfo.InvariantCulture;

        Apply(nameof(extended.Template), extended.Template,
            value => properties.Template = Element<XP.Template>(value));
        Apply(nameof(extended.Company), extended.Company,
            value => properties.Company = Element<XP.Company>(value));
        Apply(nameof(extended.Manager), extended.Manager,
            value => properties.Manager = Element<XP.Manager>(value));
        Apply(nameof(extended.Application), extended.Application,
            value => properties.Application = Element<XP.Application>(value));
        Apply(nameof(extended.AppVersion), extended.AppVersion,
            value => properties.ApplicationVersion = Element<XP.ApplicationVersion>(value));

        Apply(nameof(extended.Pages), extended.Pages?.ToString(culture),
            value => properties.Pages = Element<XP.Pages>(value));
        Apply(nameof(extended.Words), extended.Words?.ToString(culture),
            value => properties.Words = Element<XP.Words>(value));
        Apply(nameof(extended.Characters), extended.Characters?.ToString(culture),
            value => properties.Characters = Element<XP.Characters>(value));
        Apply(nameof(extended.CharactersWithSpaces), extended.CharactersWithSpaces?.ToString(culture),
            value => properties.CharactersWithSpaces = Element<XP.CharactersWithSpaces>(value));
        Apply(nameof(extended.Lines), extended.Lines?.ToString(culture),
            value => properties.Lines = Element<XP.Lines>(value));
        Apply(nameof(extended.Paragraphs), extended.Paragraphs?.ToString(culture),
            value => properties.Paragraphs = Element<XP.Paragraphs>(value));
        Apply(nameof(extended.TotalTime), extended.TotalTime?.ToString(culture),
            value => properties.TotalTime = Element<XP.TotalTime>(value));

        properties.Save();

        void Apply(string propertyName, string? value, Action<string?> assign)
        {
            if (extended.IsChanged(propertyName) || (writeAll && value is not null))
            {
                assign(value);
            }
        }
    }

    /// <summary>
    /// Builds an extended-property element, or null to remove it. Assigning null to one of the
    /// typed properties on <c>app.xml</c> removes the corresponding child element.
    /// </summary>
    private static T? Element<T>(string? value) where T : OpenXmlLeafTextElement, new() =>
        value is null ? null : new T { Text = value };

    private void ApplyCustomProperties()
    {
        var xml = _packageData.CustomPropertiesXml;
        var baseline = _packageData.Baseline?.CustomPropertiesXml;

        // Skipping unchanged properties is only valid when editing a copy of the source package,
        // which already holds them with their original types. A newly created package holds nothing,
        // so there the properties must always be written.
        if (_packageData.CanEditInPlace && xml == baseline) return;

        try
        {
            if (string.IsNullOrEmpty(xml))
            {
                if (_package!.CustomFilePropertiesPart is { } existing)
                {
                    _package.DeletePart(existing);
                }
                return;
            }

            var part = _package!.CustomFilePropertiesPart ?? _package.AddCustomFilePropertiesPart();
            WriteXmlToPart(part, xml);
        }
        catch (Exception ex)
        {
            recovery.Report("docProps/custom.xml", "Custom properties could not be written.", ex);
        }
    }

    #endregion
}
