using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Tables;
using WpTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WpTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = WordDocumentParser.Models.Tables.TableRow;
using TableCell = WordDocumentParser.Models.Tables.TableCell;

namespace WordDocumentParser.Parsing.Extractors;

/// <summary>
/// Extracts table data and formatting from Word documents.
/// </summary>
internal sealed class TableExtractor
{
    private readonly ParsingContext _context;
    private readonly Func<Paragraph, DocumentNode?> _processParagraph;
    private readonly Func<Table, DocumentNode>? _processNestedTable;

    public TableExtractor(ParsingContext context, Func<Paragraph, DocumentNode?> processParagraph, Func<Table, DocumentNode>? processNestedTable = null)
    {
        _context = context;
        _processParagraph = processParagraph;
        _processNestedTable = processNestedTable ?? ProcessTable;
    }

    /// <summary>
    /// Processes a table element and returns a DocumentNode.
    /// </summary>
    public DocumentNode ProcessTable(Table table)
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

                // Process cell content (paragraphs and nested tables)
                foreach (var element in cell.ChildElements)
                {
                    if (element is Paragraph para)
                    {
                        var paraNode = _processParagraph(para);
                        if (paraNode is not null)
                        {
                            tableCell.Content.Add(paraNode);
                        }
                    }
                    else if (element is Table nestedTable && _processNestedTable is not null)
                    {
                        var nestedTableNode = _processNestedTable(nestedTable);
                        tableCell.Content.Add(nestedTableNode);
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
        node.OriginalXml = table.OuterXml;

        return node;
    }

    /// <summary>
    /// Extracts table-level formatting properties.
    /// </summary>
    public static TableFormatting ExtractTableFormatting(Table table)
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
            formatting.TopBorder = FormattingExtractor.ExtractBorderFormatting(borders.TopBorder);
            formatting.BottomBorder = FormattingExtractor.ExtractBorderFormatting(borders.BottomBorder);
            formatting.LeftBorder = FormattingExtractor.ExtractBorderFormatting(borders.LeftBorder);
            formatting.RightBorder = FormattingExtractor.ExtractBorderFormatting(borders.RightBorder);
            formatting.InsideHorizontalBorder = FormattingExtractor.ExtractBorderFormatting(borders.InsideHorizontalBorder);
            formatting.InsideVerticalBorder = FormattingExtractor.ExtractBorderFormatting(borders.InsideVerticalBorder);
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
    /// Extracts row-level formatting properties.
    /// </summary>
    public static TableRowFormatting ExtractTableRowFormatting(WpTableRow row)
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
    /// Extracts cell-level formatting properties.
    /// </summary>
    public static TableCellFormatting ExtractTableCellFormatting(WpTableCell cell)
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
            formatting.TopBorder = FormattingExtractor.ExtractBorderFormatting(borders.TopBorder);
            formatting.BottomBorder = FormattingExtractor.ExtractBorderFormatting(borders.BottomBorder);
            formatting.LeftBorder = FormattingExtractor.ExtractBorderFormatting(borders.LeftBorder);
            formatting.RightBorder = FormattingExtractor.ExtractBorderFormatting(borders.RightBorder);
        }

        // Text direction
        formatting.TextDirection = tcPr.TextDirection?.Val?.Value.ToString();

        // No wrap
        formatting.NoWrap = tcPr.NoWrap is not null;

        return formatting;
    }
}
