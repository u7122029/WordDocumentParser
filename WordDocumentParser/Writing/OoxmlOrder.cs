using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordDocumentParser.Writing;

/// <summary>
/// Inserts property elements at their schema-mandated position.
/// </summary>
/// <remarks>
/// The OOXML schema defines <c>rPr</c>, <c>pPr</c>, and <c>tcPr</c> as ordered sequences, and Word
/// rejects a document whose children appear out of order. Appending is therefore only safe when the
/// new element sorts last, so every insertion goes through these tables.
/// </remarks>
internal static class OoxmlOrder
{
    /// <summary>Child order of <c>w:rPr</c>.</summary>
    private static readonly Type[] RunPropertyOrder =
    [
        typeof(RunStyle), typeof(RunFonts), typeof(Bold), typeof(BoldComplexScript),
        typeof(Italic), typeof(ItalicComplexScript), typeof(Caps), typeof(SmallCaps),
        typeof(Strike), typeof(DoubleStrike), typeof(Outline), typeof(Shadow),
        typeof(Emboss), typeof(Imprint), typeof(NoProof), typeof(SnapToGrid),
        typeof(Vanish), typeof(WebHidden), typeof(Color), typeof(Spacing),
        typeof(CharacterScale), typeof(Kern), typeof(Position), typeof(FontSize),
        typeof(FontSizeComplexScript), typeof(Highlight), typeof(Underline), typeof(TextEffect),
        typeof(Border), typeof(Shading), typeof(FitText), typeof(VerticalTextAlignment),
        typeof(RightToLeftText), typeof(ComplexScript), typeof(Emphasis), typeof(Languages),
        typeof(EastAsianLayout), typeof(SpecVanish)
    ];

    /// <summary>Child order of <c>w:pPr</c>.</summary>
    private static readonly Type[] ParagraphPropertyOrder =
    [
        typeof(ParagraphStyleId), typeof(KeepNext), typeof(KeepLines), typeof(PageBreakBefore),
        typeof(FrameProperties), typeof(WidowControl), typeof(NumberingProperties),
        typeof(SuppressLineNumbers), typeof(ParagraphBorders), typeof(Shading), typeof(Tabs),
        typeof(SuppressAutoHyphens), typeof(Kinsoku), typeof(WordWrap),
        typeof(OverflowPunctuation), typeof(TopLinePunctuation), typeof(AutoSpaceDE),
        typeof(AutoSpaceDN), typeof(BiDi), typeof(AdjustRightIndent), typeof(SnapToGrid),
        typeof(SpacingBetweenLines), typeof(Indentation), typeof(ContextualSpacing),
        typeof(MirrorIndents), typeof(SuppressOverlap), typeof(Justification),
        typeof(TextDirection), typeof(TextAlignment), typeof(TextBoxTightWrap),
        typeof(OutlineLevel), typeof(DivId), typeof(ConditionalFormatStyle),
        typeof(ParagraphMarkRunProperties), typeof(SectionProperties)
    ];

    /// <summary>Child order of <c>w:tcPr</c>.</summary>
    private static readonly Type[] TableCellPropertyOrder =
    [
        typeof(ConditionalFormatStyle), typeof(TableCellWidth), typeof(GridSpan),
        typeof(HorizontalMerge), typeof(VerticalMerge), typeof(TableCellBorders),
        typeof(Shading), typeof(NoWrap), typeof(TableCellMargin), typeof(TextDirection),
        typeof(TableCellFitText), typeof(TableCellVerticalAlignment), typeof(HideMark)
    ];

    /// <summary>Child order of <c>w:tblPr</c>.</summary>
    private static readonly Type[] TablePropertyOrder =
    [
        typeof(TableStyle), typeof(TablePositionProperties), typeof(TableOverlap),
        typeof(BiDiVisual), typeof(TableWidth), typeof(TableJustification),
        typeof(TableCellSpacing), typeof(TableIndentation), typeof(TableBorders),
        typeof(Shading), typeof(TableLayout), typeof(TableCellMarginDefault),
        typeof(TableLook), typeof(TableCaption), typeof(TableDescription)
    ];

    /// <summary>Inserts a child of <c>w:rPr</c> at its schema position.</summary>
    /// <param name="runProperties">The properties element.</param>
    /// <param name="element">The child to insert.</param>
    public static void InsertRunProperty(RunProperties runProperties, OpenXmlElement element)
        => InsertInOrder(runProperties, element, RunPropertyOrder);

    /// <summary>Inserts a child of <c>w:pPr</c> at its schema position.</summary>
    /// <param name="paragraphProperties">The properties element.</param>
    /// <param name="element">The child to insert.</param>
    public static void InsertParagraphProperty(ParagraphProperties paragraphProperties, OpenXmlElement element)
        => InsertInOrder(paragraphProperties, element, ParagraphPropertyOrder);

    /// <summary>Inserts a child of <c>w:tcPr</c> at its schema position.</summary>
    /// <param name="cellProperties">The properties element.</param>
    /// <param name="element">The child to insert.</param>
    public static void InsertTableCellProperty(TableCellProperties cellProperties, OpenXmlElement element)
        => InsertInOrder(cellProperties, element, TableCellPropertyOrder);

    /// <summary>Inserts a child of <c>w:tblPr</c> at its schema position.</summary>
    /// <param name="tableProperties">The properties element.</param>
    /// <param name="element">The child to insert.</param>
    public static void InsertTableProperty(TableProperties tableProperties, OpenXmlElement element)
        => InsertInOrder(tableProperties, element, TablePropertyOrder);

    private static void InsertInOrder(OpenXmlElement parent, OpenXmlElement element, Type[] order)
    {
        var position = Array.IndexOf(order, element.GetType());
        if (position < 0)
        {
            parent.Append(element);
            return;
        }

        foreach (var child in parent.ChildElements)
        {
            var childPosition = Array.IndexOf(order, child.GetType());
            if (childPosition > position)
            {
                child.InsertBeforeSelf(element);
                return;
            }
        }

        parent.Append(element);
    }
}
