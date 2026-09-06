using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using ModelTableCell = WordDocumentParser.Models.Tables.TableCell;
using ModelTableData = WordDocumentParser.Models.Tables.TableData;
using ModelTableRow = WordDocumentParser.Models.Tables.TableRow;

namespace WordDocumentParser.Writing;

/// <summary>
/// Applies a table model's pending edits onto the table XML it was parsed from.
/// </summary>
/// <remarks>
/// Only edits the caller actually made are applied. Inferring them from the presence of parsed
/// formatting would rewrite an untouched table, and ignoring deletions would leave cleared text in
/// the saved document, which matters when the caller was redacting.
/// </remarks>
internal sealed class TableEditor(Func<string, string> transformXml)
{
    /// <summary>
    /// Applies the model's edits to a table element.
    /// </summary>
    /// <param name="table">The table XML to edit in place.</param>
    /// <param name="data">The table model.</param>
    public void Apply(Table table, ModelTableData data)
    {
        if (!data.HasTableChanges) return;

        if (data.Formatting?.HasChanges is true)
        {
            ApplyTableFormatting(table, data);
        }

        var xmlRows = table.Elements<TableRow>().ToList();

        for (var rowIndex = 0; rowIndex < data.Rows.Count && rowIndex < xmlRows.Count; rowIndex++)
        {
            ApplyRow(xmlRows[rowIndex], data.Rows[rowIndex]);
        }
    }

    private static void ApplyTableFormatting(Table table, ModelTableData data)
    {
        var formatting = data.Formatting!;
        var props = table.GetFirstChild<TableProperties>();
        if (props is null)
        {
            props = new TableProperties();
            table.InsertAt(props, 0);
        }

        if (formatting.IsChanged(nameof(formatting.Alignment)))
        {
            props.GetFirstChild<TableJustification>()?.Remove();

            if (OoxmlEnum.Parse<TableRowAlignmentValues>(formatting.Alignment) is { } alignment)
            {
                OoxmlOrder.InsertTableProperty(props, new TableJustification { Val = alignment });
            }
        }

        if (formatting.IsAnyChanged(nameof(formatting.Width), nameof(formatting.WidthType)) &&
            !string.IsNullOrEmpty(formatting.Width))
        {
            props.GetFirstChild<TableWidth>()?.Remove();
            OoxmlOrder.InsertTableProperty(props, new TableWidth
            {
                Width = formatting.Width,
                Type = OoxmlEnum.Parse<TableWidthUnitValues>(formatting.WidthType) ??
                       new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa)
            });
        }

        if (formatting.IsAnyChanged(
                nameof(formatting.TopBorder), nameof(formatting.BottomBorder),
                nameof(formatting.LeftBorder), nameof(formatting.RightBorder),
                nameof(formatting.InsideHorizontalBorder), nameof(formatting.InsideVerticalBorder)) ||
            formatting.TopBorder?.HasChanges is true || formatting.BottomBorder?.HasChanges is true ||
            formatting.LeftBorder?.HasChanges is true || formatting.RightBorder?.HasChanges is true ||
            formatting.InsideHorizontalBorder?.HasChanges is true || formatting.InsideVerticalBorder?.HasChanges is true)
        {
            props.GetFirstChild<TableBorders>()?.Remove();

            var borders = new TableBorders();
            // tblBorders requires top, left, bottom, right, insideH, insideV in that order.
            if (BorderBuilder.Create<TopBorder>(formatting.TopBorder) is { } top) borders.Append(top);
            if (BorderBuilder.Create<LeftBorder>(formatting.LeftBorder) is { } left) borders.Append(left);
            if (BorderBuilder.Create<BottomBorder>(formatting.BottomBorder) is { } bottom) borders.Append(bottom);
            if (BorderBuilder.Create<RightBorder>(formatting.RightBorder) is { } right) borders.Append(right);
            if (BorderBuilder.Create<InsideHorizontalBorder>(formatting.InsideHorizontalBorder) is { } insideH) borders.Append(insideH);
            if (BorderBuilder.Create<InsideVerticalBorder>(formatting.InsideVerticalBorder) is { } insideV) borders.Append(insideV);

            if (borders.HasChildren)
            {
                OoxmlOrder.InsertTableProperty(props, borders);
            }
        }
    }

    private void ApplyRow(TableRow xmlRow, ModelTableRow dataRow)
    {
        if (!dataRow.HasRowChanges) return;

        if (dataRow.Formatting?.IsChanged(nameof(dataRow.Formatting.IsHeader)) is true)
        {
            var rowProps = xmlRow.GetFirstChild<TableRowProperties>();
            if (rowProps is null)
            {
                rowProps = new TableRowProperties();
                xmlRow.InsertAt(rowProps, 0);
            }

            rowProps.GetFirstChild<TableHeader>()?.Remove();
            if (dataRow.Formatting.IsHeader)
            {
                rowProps.Append(new TableHeader());
            }
        }

        var xmlCells = xmlRow.Elements<TableCell>().ToList();
        for (var cellIndex = 0; cellIndex < dataRow.Cells.Count && cellIndex < xmlCells.Count; cellIndex++)
        {
            ApplyCell(xmlCells[cellIndex], dataRow.Cells[cellIndex]);
        }
    }

    private void ApplyCell(TableCell xmlCell, ModelTableCell dataCell)
    {
        if (!dataCell.HasCellChanges) return;

        if (dataCell.Formatting is { } formatting && (formatting.HasChanges || formatting.HasBorderChanges))
        {
            ApplyCellFormatting(xmlCell, dataCell);
        }

        ApplyCellContent(xmlCell, dataCell);
    }

    private static void ApplyCellFormatting(TableCell xmlCell, ModelTableCell dataCell)
    {
        var formatting = dataCell.Formatting!;
        var cellProps = xmlCell.GetFirstChild<TableCellProperties>();
        if (cellProps is null)
        {
            cellProps = new TableCellProperties();
            xmlCell.InsertAt(cellProps, 0);
        }

        if (formatting.IsAnyChanged(
                nameof(formatting.ShadingFill), nameof(formatting.ShadingColor), nameof(formatting.ShadingPattern)))
        {
            cellProps.GetFirstChild<Shading>()?.Remove();

            if (!string.IsNullOrEmpty(formatting.ShadingFill))
            {
                var shading = new Shading
                {
                    Fill = formatting.ShadingFill,
                    Val = OoxmlEnum.Parse<ShadingPatternValues>(formatting.ShadingPattern) ??
                          new EnumValue<ShadingPatternValues>(ShadingPatternValues.Clear)
                };
                if (!string.IsNullOrEmpty(formatting.ShadingColor))
                {
                    shading.Color = formatting.ShadingColor;
                }
                OoxmlOrder.InsertTableCellProperty(cellProps, shading);
            }
        }

        if (formatting.IsChanged(nameof(formatting.VerticalAlignment)))
        {
            cellProps.GetFirstChild<TableCellVerticalAlignment>()?.Remove();

            if (OoxmlEnum.Parse<TableVerticalAlignmentValues>(formatting.VerticalAlignment) is { } alignment)
            {
                OoxmlOrder.InsertTableCellProperty(cellProps, new TableCellVerticalAlignment { Val = alignment });
            }
        }

        if (formatting.IsChanged(nameof(formatting.NoWrap)))
        {
            cellProps.GetFirstChild<NoWrap>()?.Remove();
            if (formatting.NoWrap)
            {
                OoxmlOrder.InsertTableCellProperty(cellProps, new NoWrap());
            }
        }

        // Borders are rewritten only on an explicit border edit: the schema fixes their order, so
        // regenerating them speculatively risks reordering valid XML into invalid XML.
        if (formatting.HasBorderChanges)
        {
            cellProps.GetFirstChild<TableCellBorders>()?.Remove();

            var borders = new TableCellBorders();
            // tcBorders requires top, left, bottom, right in that order.
            if (BorderBuilder.Create<TopBorder>(formatting.TopBorder) is { } top) borders.Append(top);
            if (BorderBuilder.Create<LeftBorder>(formatting.LeftBorder) is { } left) borders.Append(left);
            if (BorderBuilder.Create<BottomBorder>(formatting.BottomBorder) is { } bottom) borders.Append(bottom);
            if (BorderBuilder.Create<RightBorder>(formatting.RightBorder) is { } right) borders.Append(right);

            if (borders.HasChildren)
            {
                OoxmlOrder.InsertTableCellProperty(cellProps, borders);
            }
        }
    }

    /// <summary>
    /// Reconciles a cell's content nodes with the paragraphs and nested tables in its XML.
    /// </summary>
    private void ApplyCellContent(TableCell xmlCell, ModelTableCell dataCell)
    {
        if (dataCell.IsContentChanged)
        {
            RebuildCellContent(xmlCell, dataCell);
            return;
        }

        var xmlParagraphs = xmlCell.Elements<Paragraph>().ToList();
        var xmlTables = xmlCell.Elements<Table>().ToList();
        var paragraphIndex = 0;
        var tableIndex = 0;

        foreach (var contentNode in dataCell.Content)
        {
            if (contentNode.Type == ContentType.Table)
            {
                if (tableIndex < xmlTables.Count)
                {
                    ApplyNestedTable(xmlTables[tableIndex], contentNode);
                }
                tableIndex++;
                continue;
            }

            if (paragraphIndex < xmlParagraphs.Count)
            {
                if (contentNode.HasChanges)
                {
                    ParagraphEditor.Apply(xmlParagraphs[paragraphIndex], contentNode);
                }
            }
            paragraphIndex++;
        }
    }

    /// <summary>
    /// Replaces a nested table with the node's current XML, then applies its model edits.
    /// </summary>
    /// <remarks>
    /// Structural edits — adding a row, inserting a column — rewrite the nested node's own
    /// <c>OriginalXml</c>, which the enclosing table's XML knows nothing about. Reading it back here
    /// is what makes those edits reach the saved document.
    /// </remarks>
    private void ApplyNestedTable(Table xmlTable, DocumentNode node)
    {
        var replacement = xmlTable;

        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            replacement = new Table(transformXml(node.OriginalXml));
            xmlTable.InsertAfterSelf(replacement);
            xmlTable.Remove();
        }

        if (node.GetTableData() is { } nestedData)
        {
            Apply(replacement, nestedData);
        }
    }

    /// <summary>
    /// Rebuilds a cell's body to match its content nodes, keeping the cell's own properties.
    /// </summary>
    /// <remarks>
    /// A cell must contain at least one paragraph, so emptying one leaves a single empty paragraph
    /// rather than no content at all.
    /// </remarks>
    private void RebuildCellContent(TableCell xmlCell, ModelTableCell dataCell)
    {
        var template = xmlCell.Elements<Paragraph>().FirstOrDefault();
        var templateProps = template?.ParagraphProperties;
        var templateRunProps = template?.Elements<Run>().FirstOrDefault()?.RunProperties;

        foreach (var child in xmlCell.ChildElements.ToList())
        {
            if (child is not TableCellProperties)
            {
                child.Remove();
            }
        }

        foreach (var contentNode in dataCell.Content)
        {
            if (contentNode.Type == ContentType.Table && !string.IsNullOrEmpty(contentNode.OriginalXml))
            {
                var nested = new Table(transformXml(contentNode.OriginalXml));
                if (contentNode.GetTableData() is { } nestedData)
                {
                    Apply(nested, nestedData);
                }
                xmlCell.Append(nested);
                continue;
            }

            xmlCell.Append(BuildParagraph(contentNode, templateProps, templateRunProps));
        }

        if (!xmlCell.Elements<Paragraph>().Any())
        {
            xmlCell.Append(new Paragraph());
        }
    }

    private static Paragraph BuildParagraph(
        DocumentNode node, ParagraphProperties? templateProps, RunProperties? templateRunProps)
    {
        // An unedited node still holding its source XML is reproduced exactly.
        if (!node.HasChanges && !string.IsNullOrEmpty(node.OriginalXml))
        {
            return new Paragraph(node.OriginalXml);
        }

        var paragraph = new Paragraph();
        if (templateProps is not null)
        {
            paragraph.Append((ParagraphProperties)templateProps.CloneNode(true));
        }

        if (node.HasFormattedRuns)
        {
            foreach (var modelRun in node.Runs)
            {
                paragraph.Append(BuildRun(modelRun, templateRunProps));
            }
        }
        else if (!string.IsNullOrEmpty(node.Text))
        {
            var run = new Run();
            if (templateRunProps is not null)
            {
                run.RunProperties = (RunProperties)templateRunProps.CloneNode(true);
            }
            run.Append(new Text(node.Text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);
        }

        if (node.ParagraphFormatting?.HasFormattingChanges is true)
        {
            ParagraphEditor.ApplyParagraphFormatting(paragraph, node.ParagraphFormatting);
        }

        return paragraph;
    }

    private static Run BuildRun(Models.Formatting.FormattedRun modelRun, RunProperties? templateRunProps)
    {
        var run = new Run();
        if (templateRunProps is not null)
        {
            run.RunProperties = (RunProperties)templateRunProps.CloneNode(true);
        }

        if (modelRun.Formatting.HasChanges)
        {
            RunEditor.ApplyRunFormatting(run, modelRun.Formatting);
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
}
