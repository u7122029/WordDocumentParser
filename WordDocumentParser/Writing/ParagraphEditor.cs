using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Models.ContentControls;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Writing;

/// <summary>
/// Applies a node's pending edits onto the paragraph or SDT it was parsed from.
/// </summary>
/// <remarks>
/// The element keeps its original identity: only the properties a caller actually assigned are
/// touched. An unedited paragraph is therefore reproduced as it was, and a font change reaches the
/// text without taking the surrounding hyperlink, field, or bookmark with it.
/// </remarks>
internal static class ParagraphEditor
{
    /// <summary>
    /// Applies every pending edit on <paramref name="node"/> to <paramref name="element"/>.
    /// </summary>
    /// <param name="element">A <c>w:p</c> or <c>w:sdt</c> parsed from the node's original XML.</param>
    /// <param name="node">The node holding the edits.</param>
    public static void Apply(OpenXmlElement element, DocumentNode node)
    {
        if (!node.HasChanges) return;

        switch (element)
        {
            case Paragraph paragraph:
                ApplyToParagraph(paragraph, node);
                break;

            case SdtBlock sdtBlock:
                ApplyToSdtBlock(sdtBlock, node);
                break;
        }
    }

    private static void ApplyToParagraph(Paragraph paragraph, DocumentNode node)
    {
        if (node.IsParagraphFormattingChanged && node.ParagraphFormatting is not null)
        {
            ApplyParagraphFormatting(paragraph, node.ParagraphFormatting);
        }

        // An inline control's own definition — its checkbox state, tag, date — lives in the SDT
        // properties, which the general run pass knows nothing about.
        ApplyInlineContentControlProperties(paragraph, node);

        RunEditor.ApplyRunEdits(paragraph, node);
    }

    private static void ApplyToSdtBlock(SdtBlock sdtBlock, DocumentNode node)
    {
        if (node.IsContentControlChanged && node.ContentControlProperties is not null &&
            sdtBlock.SdtProperties is not null)
        {
            ApplyContentControlProperties(sdtBlock.SdtProperties, node.ContentControlProperties);
        }

        var content = sdtBlock.SdtContentBlock;
        if (content is null) return;

        // A single-paragraph control: edit that paragraph in place so field codes and formatting
        // inside the control survive.
        var paragraphs = content.Elements<Paragraph>().ToList();
        if (paragraphs.Count == 1)
        {
            ApplyToParagraph(paragraphs[0], node);
            return;
        }

        // Multi-element control with an assigned text value and nowhere obvious to put it: replace
        // the first paragraph's text and leave the rest of the control alone.
        if (node.IsTextChanged && paragraphs.Count > 0)
        {
            RunEditor.ReplaceParagraphText(paragraphs[0], node.Text);
        }
    }

    /// <summary>
    /// Applies changed properties of the paragraph's inline content controls to their
    /// <c>w:sdtPr</c> elements.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Only the control's definition is touched here. Its text and formatting go through the general
    /// run pass, which maps model runs to the runs that actually hold them. Consolidating a control's
    /// text into its first run instead would give every part of it that run's formatting — bold text
    /// followed by italic text came back bold twice.
    /// </para>
    /// <para>
    /// Nothing here writes back to the model: serialization must not consume the edits it is
    /// serializing, or a second save of the same instance would emit the pre-edit document.
    /// </para>
    /// </remarks>
    private static void ApplyInlineContentControlProperties(Paragraph paragraph, DocumentNode node)
    {
        var changed = new Dictionary<int, ContentControlProperties>();

        foreach (var run in node.Runs)
        {
            var props = run.ContentControlProperties;
            if (props?.Id is { } id && props.HasChanges)
            {
                changed[id] = props;
            }
        }

        if (changed.Count == 0) return;

        foreach (var sdtRun in paragraph.Descendants<SdtRun>().ToList())
        {
            if (sdtRun.SdtProperties is not { } sdtProperties) continue;
            if (sdtProperties.GetFirstChild<SdtId>()?.Val?.Value is not { } id) continue;

            if (changed.TryGetValue(id, out var props))
            {
                ApplyContentControlProperties(sdtProperties, props);
            }
        }
    }

    /// <summary>
    /// Applies only the changed properties of a content control to its <c>w:sdtPr</c>, leaving the
    /// rest of the definition — including extension-namespace children such as
    /// <c>w14:checkbox</c> — untouched.
    /// </summary>
    public static void ApplyContentControlProperties(SdtProperties sdtProperties, ContentControlProperties props)
    {
        if (props.IsChanged(nameof(ContentControlProperties.Tag)))
        {
            SetOrRemoveChild(sdtProperties, props.Tag, value => new Tag { Val = value });
        }

        if (props.IsChanged(nameof(ContentControlProperties.Alias)))
        {
            SetOrRemoveChild(sdtProperties, props.Alias, value => new SdtAlias { Val = value });
        }

        if (props.IsChanged(nameof(ContentControlProperties.IsChecked)))
        {
            ApplyCheckboxState(sdtProperties, props.IsChecked);
        }

        if (props.IsChanged(nameof(ContentControlProperties.DateValue)) && props.DateValue.HasValue)
        {
            var date = sdtProperties.GetFirstChild<SdtContentDate>();
            if (date is not null)
            {
                date.FullDate = props.DateValue.Value;
            }
        }
    }

    /// <summary>
    /// Updates the <c>w14:checked</c> value of a checkbox control, preserving the rest of the
    /// element and its attributes.
    /// </summary>
    private static void ApplyCheckboxState(SdtProperties sdtProperties, bool? isChecked)
    {
        var checkbox = sdtProperties.Descendants().FirstOrDefault(e => e.LocalName == "checkbox");
        var checkedElement = checkbox?.Descendants().FirstOrDefault(e => e.LocalName == "checked");
        if (checkedElement is null) return;

        var newValue = isChecked == true ? "1" : "0";
        var attributes = checkedElement.GetAttributes().ToList();
        var valueAttribute = attributes.FindIndex(a => a.LocalName == "val");

        if (valueAttribute < 0) return;

        var existing = attributes[valueAttribute];
        checkedElement.SetAttribute(
            new OpenXmlAttribute(existing.Prefix, existing.LocalName, existing.NamespaceUri, newValue));
    }

    /// <summary>
    /// Applies only the changed paragraph properties onto the paragraph's existing <c>w:pPr</c>.
    /// </summary>
    public static void ApplyParagraphFormatting(Paragraph paragraph, ParagraphFormatting formatting)
    {
        var props = paragraph.ParagraphProperties;
        if (props is null)
        {
            props = new ParagraphProperties();
            paragraph.InsertAt(props, 0);
        }

        if (formatting.IsChanged(nameof(ParagraphFormatting.StyleId)))
        {
            SetOrRemoveChild(props, formatting.StyleId, value => new ParagraphStyleId { Val = value },
                OoxmlOrder.InsertParagraphProperty);
        }

        if (formatting.IsChanged(nameof(ParagraphFormatting.Alignment)))
        {
            SetOrRemoveChild(props, formatting.Alignment,
                value => OoxmlEnum.Parse<JustificationValues>(value) is { } parsed
                    ? new Justification { Val = parsed }
                    : null,
                OoxmlOrder.InsertParagraphProperty);
        }

        SetToggle<KeepNext>(props, formatting, nameof(ParagraphFormatting.KeepNext), formatting.KeepNext);
        SetToggle<KeepLines>(props, formatting, nameof(ParagraphFormatting.KeepLines), formatting.KeepLines);
        SetToggle<PageBreakBefore>(props, formatting, nameof(ParagraphFormatting.PageBreakBefore), formatting.PageBreakBefore);
        SetToggle<WidowControl>(props, formatting, nameof(ParagraphFormatting.WidowControl), formatting.WidowControl);

        if (formatting.IsAnyChanged(
                nameof(ParagraphFormatting.IndentLeft), nameof(ParagraphFormatting.IndentRight),
                nameof(ParagraphFormatting.IndentFirstLine), nameof(ParagraphFormatting.IndentHanging)))
        {
            var indentation = props.GetFirstChild<Indentation>();
            if (indentation is null)
            {
                indentation = new Indentation();
                OoxmlOrder.InsertParagraphProperty(props, indentation);
            }

            if (formatting.IsChanged(nameof(ParagraphFormatting.IndentLeft))) indentation.Left = formatting.IndentLeft;
            if (formatting.IsChanged(nameof(ParagraphFormatting.IndentRight))) indentation.Right = formatting.IndentRight;
            if (formatting.IsChanged(nameof(ParagraphFormatting.IndentFirstLine))) indentation.FirstLine = formatting.IndentFirstLine;
            if (formatting.IsChanged(nameof(ParagraphFormatting.IndentHanging))) indentation.Hanging = formatting.IndentHanging;
        }

        if (formatting.IsAnyChanged(
                nameof(ParagraphFormatting.SpacingBefore), nameof(ParagraphFormatting.SpacingAfter),
                nameof(ParagraphFormatting.LineSpacing), nameof(ParagraphFormatting.LineSpacingRule)))
        {
            var spacing = props.GetFirstChild<SpacingBetweenLines>();
            if (spacing is null)
            {
                spacing = new SpacingBetweenLines();
                OoxmlOrder.InsertParagraphProperty(props, spacing);
            }

            if (formatting.IsChanged(nameof(ParagraphFormatting.SpacingBefore))) spacing.Before = formatting.SpacingBefore;
            if (formatting.IsChanged(nameof(ParagraphFormatting.SpacingAfter))) spacing.After = formatting.SpacingAfter;
            if (formatting.IsChanged(nameof(ParagraphFormatting.LineSpacing))) spacing.Line = formatting.LineSpacing;
            if (formatting.IsChanged(nameof(ParagraphFormatting.LineSpacingRule)))
            {
                spacing.LineRule = OoxmlEnum.Parse<LineSpacingRuleValues>(formatting.LineSpacingRule);
            }
        }

        if (formatting.IsAnyChanged(nameof(ParagraphFormatting.ShadingFill), nameof(ParagraphFormatting.ShadingColor)))
        {
            props.GetFirstChild<Shading>()?.Remove();
            if (!string.IsNullOrEmpty(formatting.ShadingFill))
            {
                OoxmlOrder.InsertParagraphProperty(props, new Shading
                {
                    Fill = formatting.ShadingFill,
                    Color = formatting.ShadingColor
                });
            }
        }

        if (formatting.IsAnyChanged(nameof(ParagraphFormatting.NumberingId), nameof(ParagraphFormatting.NumberingLevel)))
        {
            props.GetFirstChild<NumberingProperties>()?.Remove();
            if (formatting.NumberingId.HasValue)
            {
                OoxmlOrder.InsertParagraphProperty(props, new NumberingProperties(
                    new NumberingLevelReference { Val = formatting.NumberingLevel ?? 0 },
                    new NumberingId { Val = formatting.NumberingId.Value }));
            }
        }

        if (formatting.IsAnyChanged(
                nameof(ParagraphFormatting.TopBorder), nameof(ParagraphFormatting.BottomBorder),
                nameof(ParagraphFormatting.LeftBorder), nameof(ParagraphFormatting.RightBorder)) ||
            formatting.HasFormattingChanges && HasChangedBorder(formatting))
        {
            ApplyParagraphBorders(props, formatting);
        }

        if (!props.HasChildren)
        {
            props.Remove();
        }
    }

    private static bool HasChangedBorder(ParagraphFormatting formatting) =>
        formatting.TopBorder?.HasChanges is true || formatting.BottomBorder?.HasChanges is true ||
        formatting.LeftBorder?.HasChanges is true || formatting.RightBorder?.HasChanges is true;

    private static void ApplyParagraphBorders(ParagraphProperties props, ParagraphFormatting formatting)
    {
        props.GetFirstChild<ParagraphBorders>()?.Remove();

        if (formatting.TopBorder is null && formatting.BottomBorder is null &&
            formatting.LeftBorder is null && formatting.RightBorder is null)
        {
            return;
        }

        var borders = new ParagraphBorders();

        // pBdr requires top, left, bottom, right in that order.
        if (BorderBuilder.Create<TopBorder>(formatting.TopBorder) is { } top) borders.Append(top);
        if (BorderBuilder.Create<LeftBorder>(formatting.LeftBorder) is { } left) borders.Append(left);
        if (BorderBuilder.Create<BottomBorder>(formatting.BottomBorder) is { } bottom) borders.Append(bottom);
        if (BorderBuilder.Create<RightBorder>(formatting.RightBorder) is { } right) borders.Append(right);

        if (borders.HasChildren)
        {
            OoxmlOrder.InsertParagraphProperty(props, borders);
        }
    }

    private static void SetToggle<T>(
        ParagraphProperties props, ParagraphFormatting formatting, string propertyName, bool enabled)
        where T : OpenXmlLeafElement, new()
    {
        if (!formatting.IsChanged(propertyName)) return;

        props.GetFirstChild<T>()?.Remove();
        if (enabled)
        {
            OoxmlOrder.InsertParagraphProperty(props, new T());
        }
    }

    private static void SetOrRemoveChild<TParent, TChild>(
        TParent parent, string? value, Func<string, TChild?> factory, Action<TParent, OpenXmlElement>? insert = null)
        where TParent : OpenXmlElement
        where TChild : OpenXmlElement
    {
        parent.GetFirstChild<TChild>()?.Remove();
        if (string.IsNullOrEmpty(value)) return;

        var element = factory(value);
        if (element is null) return;

        if (insert is not null)
        {
            insert(parent, element);
        }
        else
        {
            parent.Append(element);
        }
    }

    private static void SetOrRemoveChild<TChild>(SdtProperties parent, string? value, Func<string, TChild?> factory)
        where TChild : OpenXmlElement
        => SetOrRemoveChild(parent, value, factory, null);
}

/// <summary>
/// Builds typed border elements from the model's border formatting.
/// </summary>
internal static class BorderBuilder
{
    /// <summary>
    /// Creates a border element, or null when there is no formatting to write.
    /// </summary>
    /// <typeparam name="T">The border element type.</typeparam>
    /// <param name="formatting">The border formatting, which may be null.</param>
    /// <returns>The element, or null.</returns>
    public static T? Create<T>(BorderFormatting? formatting) where T : BorderType, new()
    {
        if (formatting is null) return null;

        var border = new T
        {
            // A border with no explicit style is a single line, matching what Word writes.
            Val = OoxmlEnum.Parse<BorderValues>(formatting.Style) ??
                  new EnumValue<BorderValues>(BorderValues.Single)
        };

        if (uint.TryParse(formatting.Size, out var size)) border.Size = size;
        if (!string.IsNullOrEmpty(formatting.Color)) border.Color = formatting.Color;
        if (uint.TryParse(formatting.Space, out var space)) border.Space = space;

        return border;
    }
}
