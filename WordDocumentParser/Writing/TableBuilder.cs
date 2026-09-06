using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using ModelTableCell = WordDocumentParser.Models.Tables.TableCell;
using ModelTableData = WordDocumentParser.Models.Tables.TableData;
using ModelTableRow = WordDocumentParser.Models.Tables.TableRow;

namespace WordDocumentParser.Writing;

/// <summary>
/// Builds table XML from a table model, for tables that were created in code rather than parsed.
/// </summary>
internal static class TableBuilder
{
    /// <summary>Builds a complete table element.</summary>
    /// <param name="data">The table model.</param>
    /// <returns>The table element.</returns>
    public static Table Build(ModelTableData data)
    {
        var table = new Table();
        table.Append(BuildTableProperties(data.Formatting));
        table.Append(BuildGrid(data));

        foreach (var row in data.Rows)
        {
            table.Append(BuildRow(row));
        }

        return table;
    }

    private static TableProperties BuildTableProperties(Models.Formatting.TableFormatting? formatting)
    {
        var properties = new TableProperties();

        properties.Append(string.IsNullOrEmpty(formatting?.Width)
            ? new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
            : new TableWidth
            {
                Width = formatting.Width,
                Type = OoxmlEnum.Parse<TableWidthUnitValues>(formatting.WidthType) ??
                       new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa)
            });

        if (OoxmlEnum.Parse<TableRowAlignmentValues>(formatting?.Alignment) is { } alignment)
        {
            properties.Append(new TableJustification { Val = alignment });
        }

        var borders = new TableBorders();
        // tblBorders requires top, left, bottom, right, insideH, insideV in that order.
        if (BorderBuilder.Create<TopBorder>(formatting?.TopBorder) is { } top) borders.Append(top);
        if (BorderBuilder.Create<LeftBorder>(formatting?.LeftBorder) is { } left) borders.Append(left);
        if (BorderBuilder.Create<BottomBorder>(formatting?.BottomBorder) is { } bottom) borders.Append(bottom);
        if (BorderBuilder.Create<RightBorder>(formatting?.RightBorder) is { } right) borders.Append(right);
        if (BorderBuilder.Create<InsideHorizontalBorder>(formatting?.InsideHorizontalBorder) is { } insideH) borders.Append(insideH);
        if (BorderBuilder.Create<InsideVerticalBorder>(formatting?.InsideVerticalBorder) is { } insideV) borders.Append(insideV);

        properties.Append(borders.HasChildren
            ? borders
            : new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));

        return properties;
    }

    private static TableGrid BuildGrid(ModelTableData data)
    {
        var grid = new TableGrid();
        var widths = data.Formatting?.GridColumnWidths;

        if (widths is { Count: > 0 })
        {
            foreach (var width in widths)
            {
                var column = new GridColumn();
                if (!string.IsNullOrEmpty(width)) column.Width = width;
                grid.Append(column);
            }
        }
        else
        {
            for (var i = 0; i < data.ColumnCount; i++)
            {
                grid.Append(new GridColumn());
            }
        }

        return grid;
    }

    private static TableRow BuildRow(ModelTableRow row)
    {
        var xmlRow = new TableRow();
        var properties = new TableRowProperties();

        var formatting = row.Formatting;
        if (formatting is not null)
        {
            if (!string.IsNullOrEmpty(formatting.Height) && uint.TryParse(formatting.Height, out var height))
            {
                properties.Append(new TableRowHeight
                {
                    Val = height,
                    HeightType = OoxmlEnum.Parse<HeightRuleValues>(formatting.HeightRule) ??
                                 new EnumValue<HeightRuleValues>(HeightRuleValues.Auto)
                });
            }

            if (formatting.CantSplit) properties.Append(new CantSplit());
            if (formatting.IsHeader) properties.Append(new TableHeader());
        }
        else if (row.IsHeader)
        {
            properties.Append(new TableHeader());
        }

        if (properties.HasChildren)
        {
            xmlRow.Append(properties);
        }

        foreach (var cell in row.Cells)
        {
            xmlRow.Append(BuildCell(cell));
        }

        return xmlRow;
    }

    private static TableCell BuildCell(ModelTableCell cell)
    {
        var xmlCell = new TableCell();
        xmlCell.Append(BuildCellProperties(cell.Formatting));

        foreach (var content in cell.Content)
        {
            if (content.Type == ContentType.Table && content.GetTableData() is { } nested)
            {
                xmlCell.Append(Build(nested));
                continue;
            }

            xmlCell.Append(BuildCellParagraph(content));
        }

        // A cell must contain at least one block-level element.
        if (!xmlCell.Elements<Paragraph>().Any() && !xmlCell.Elements<Table>().Any())
        {
            xmlCell.Append(new Paragraph());
        }
        else if (xmlCell.Elements<Table>().Any() && xmlCell.ChildElements.Last() is Table)
        {
            // A cell whose last child is a table needs a trailing paragraph.
            xmlCell.Append(new Paragraph());
        }

        return xmlCell;
    }

    private static TableCellProperties BuildCellProperties(Models.Formatting.TableCellFormatting? formatting)
    {
        var properties = new TableCellProperties();
        if (formatting is null) return properties;

        // tcPr order: tcW, gridSpan, vMerge, tcBorders, shd, noWrap, vAlign.
        if (!string.IsNullOrEmpty(formatting.Width))
        {
            properties.Append(new TableCellWidth
            {
                Width = formatting.Width,
                Type = OoxmlEnum.Parse<TableWidthUnitValues>(formatting.WidthType) ??
                       new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa)
            });
        }

        if (formatting.GridSpan > 1)
        {
            properties.Append(new GridSpan { Val = formatting.GridSpan });
        }

        if (!string.IsNullOrEmpty(formatting.VerticalMerge))
        {
            var merge = new VerticalMerge();
            if (formatting.VerticalMerge == "Restart") merge.Val = MergedCellValues.Restart;
            properties.Append(merge);
        }

        var borders = new TableCellBorders();
        if (BorderBuilder.Create<TopBorder>(formatting.TopBorder) is { } top) borders.Append(top);
        if (BorderBuilder.Create<LeftBorder>(formatting.LeftBorder) is { } left) borders.Append(left);
        if (BorderBuilder.Create<BottomBorder>(formatting.BottomBorder) is { } bottom) borders.Append(bottom);
        if (BorderBuilder.Create<RightBorder>(formatting.RightBorder) is { } right) borders.Append(right);
        if (borders.HasChildren) properties.Append(borders);

        if (!string.IsNullOrEmpty(formatting.ShadingFill))
        {
            var shading = new Shading
            {
                Fill = formatting.ShadingFill,
                Val = OoxmlEnum.Parse<ShadingPatternValues>(formatting.ShadingPattern) ??
                      new EnumValue<ShadingPatternValues>(ShadingPatternValues.Clear)
            };
            if (!string.IsNullOrEmpty(formatting.ShadingColor)) shading.Color = formatting.ShadingColor;
            properties.Append(shading);
        }

        if (formatting.NoWrap) properties.Append(new NoWrap());

        if (OoxmlEnum.Parse<TableVerticalAlignmentValues>(formatting.VerticalAlignment) is { } verticalAlignment)
        {
            properties.Append(new TableCellVerticalAlignment { Val = verticalAlignment });
        }

        return properties;
    }

    private static Paragraph BuildCellParagraph(DocumentNode node)
    {
        if (!string.IsNullOrEmpty(node.OriginalXml) && !node.HasChanges)
        {
            return new Paragraph(node.OriginalXml);
        }

        var paragraph = new Paragraph();

        if (node.ParagraphFormatting is { } formatting && formatting.HasFormatting)
        {
            ParagraphEditor.ApplyParagraphFormatting(paragraph, MarkAllParagraphProperties(formatting));
        }

        if (node.HasFormattedRuns)
        {
            foreach (var run in node.Runs)
            {
                var xmlRun = new Run();
                if (run.Formatting.HasFormatting || run.Formatting.HasChanges)
                {
                    RunEditor.ApplyRunFormatting(xmlRun, run.Formatting);
                }
                xmlRun.Append(new Text(run.Text) { Space = SpaceProcessingModeValues.Preserve });
                paragraph.Append(xmlRun);
            }
        }
        else
        {
            paragraph.Append(new Run(new Text(node.Text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        return paragraph;
    }

    /// <summary>
    /// Marks every property of a paragraph's formatting as changed, so a paragraph written from
    /// scratch emits all of it.
    /// </summary>
    private static Models.Formatting.ParagraphFormatting MarkAllParagraphProperties(
        Models.Formatting.ParagraphFormatting formatting)
    {
        var copy = formatting.Clone();
        foreach (var name in ParagraphPropertyNames)
        {
            copy.MarkChanged(name);
        }
        return copy;
    }

    private static readonly string[] ParagraphPropertyNames =
    [
        nameof(Models.Formatting.ParagraphFormatting.StyleId), nameof(Models.Formatting.ParagraphFormatting.Alignment),
        nameof(Models.Formatting.ParagraphFormatting.IndentLeft), nameof(Models.Formatting.ParagraphFormatting.IndentRight),
        nameof(Models.Formatting.ParagraphFormatting.IndentFirstLine), nameof(Models.Formatting.ParagraphFormatting.IndentHanging),
        nameof(Models.Formatting.ParagraphFormatting.SpacingBefore), nameof(Models.Formatting.ParagraphFormatting.SpacingAfter),
        nameof(Models.Formatting.ParagraphFormatting.LineSpacing), nameof(Models.Formatting.ParagraphFormatting.LineSpacingRule),
        nameof(Models.Formatting.ParagraphFormatting.KeepNext), nameof(Models.Formatting.ParagraphFormatting.KeepLines),
        nameof(Models.Formatting.ParagraphFormatting.PageBreakBefore), nameof(Models.Formatting.ParagraphFormatting.WidowControl),
        nameof(Models.Formatting.ParagraphFormatting.ShadingFill), nameof(Models.Formatting.ParagraphFormatting.ShadingColor),
        nameof(Models.Formatting.ParagraphFormatting.NumberingId), nameof(Models.Formatting.ParagraphFormatting.NumberingLevel),
        nameof(Models.Formatting.ParagraphFormatting.TopBorder), nameof(Models.Formatting.ParagraphFormatting.BottomBorder),
        nameof(Models.Formatting.ParagraphFormatting.LeftBorder), nameof(Models.Formatting.ParagraphFormatting.RightBorder)
    ];
}
