using DocumentFormat.OpenXml.Wordprocessing;

namespace WordDocumentParser.Writing;

/// <summary>
/// Builds the style and numbering definitions used by documents created in code, which have no
/// source package to inherit them from.
/// </summary>
internal static class DefaultParts
{
    private static readonly int[] HeadingSizes = [32, 26, 24, 22, 20, 18, 16, 15, 14];

    private static readonly string[] HeadingColors =
        ["2F5496", "2F5496", "1F3763", "1F3763", "1F3763", "1F3763", "1F3763", "1F3763", "1F3763"];

    /// <summary>Creates the default style definitions.</summary>
    /// <returns>The styles part content.</returns>
    public static Styles CreateStyles()
    {
        var styles = new Styles();

        var normal = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        };
        normal.Append(new StyleName { Val = "Normal" });
        normal.Append(new PrimaryStyle());
        styles.Append(normal);

        for (var level = 1; level <= 9; level++)
        {
            styles.Append(CreateHeadingStyle(level, HeadingSizes[level - 1], HeadingColors[level - 1]));
        }

        var listParagraph = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "ListParagraph"
        };
        listParagraph.Append(new StyleName { Val = "List Paragraph" });
        listParagraph.Append(new BasedOn { Val = "Normal" });
        listParagraph.Append(new StyleParagraphProperties(new Indentation { Left = "720" }));
        styles.Append(listParagraph);

        var hyperlink = new Style
        {
            Type = StyleValues.Character,
            StyleId = "Hyperlink"
        };
        hyperlink.Append(new StyleName { Val = "Hyperlink" });
        hyperlink.Append(new StyleRunProperties(
            new Color { Val = "0563C1" },
            new Underline { Val = UnderlineValues.Single }));
        styles.Append(hyperlink);

        return styles;
    }

    /// <summary>
    /// Creates one heading style.
    /// </summary>
    /// <remarks>
    /// The run properties follow the schema's <c>rPr</c> sequence — colour precedes the font sizes.
    /// Appending colour last produced nine validation errors in every document this library created
    /// from scratch.
    /// </remarks>
    private static Style CreateHeadingStyle(int level, int fontSize, string color)
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

        var paragraphProperties = new StyleParagraphProperties();
        paragraphProperties.Append(new KeepNext());
        paragraphProperties.Append(new KeepLines());
        paragraphProperties.Append(new SpacingBetweenLines { Before = level == 1 ? "240" : "160", After = "80" });
        paragraphProperties.Append(new OutlineLevel { Val = level - 1 });
        style.Append(paragraphProperties);

        var runProperties = new StyleRunProperties();
        runProperties.Append(new Bold());
        runProperties.Append(new Color { Val = color });
        runProperties.Append(new FontSize { Val = (fontSize * 2).ToString() });
        runProperties.Append(new FontSizeComplexScript { Val = (fontSize * 2).ToString() });
        style.Append(runProperties);

        return style;
    }

    /// <summary>Creates a bullet numbering definition with nine levels.</summary>
    /// <returns>The numbering part content.</returns>
    public static Numbering CreateBulletNumbering()
    {
        var numbering = new Numbering();

        var abstractNum = new AbstractNum { AbstractNumberId = 0 };
        abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        string[] bullets = ["●", "○", "■", "□", "◆", "◇", "▸", "▹", "•"];
        for (var i = 0; i < 9; i++)
        {
            var level = new Level { LevelIndex = i };
            level.Append(new StartNumberingValue { Val = 1 });
            level.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
            level.Append(new LevelText { Val = bullets[i] });
            level.Append(new LevelJustification { Val = LevelJustificationValues.Left });
            level.Append(new PreviousParagraphProperties(
                new Indentation { Left = ((i + 1) * 720).ToString(), Hanging = "360" }));
            abstractNum.Append(level);
        }

        numbering.Append(abstractNum);

        for (var i = 1; i <= 10; i++)
        {
            numbering.Append(new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = i });
        }

        return numbering;
    }
}
