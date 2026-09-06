using DocumentFormat.OpenXml;
using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Tables;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for querying and modifying tables.
/// </summary>
public static class TableExtensions
{
    #region Finding tables

    /// <summary>
    /// Gets all tables in the document, including nested tables within table cells.
    /// </summary>
    /// <param name="document">The document to search</param>
    /// <param name="includeNested">If true, includes tables nested within other table cells</param>
    /// <returns>All table nodes in the document</returns>
    public static IEnumerable<DocumentNode> FindAllTables(this WordDocument document, bool includeNested = true)
        => document.Root.FindAllTables(includeNested);

    /// <summary>
    /// Gets all tables starting from a node, including nested tables within table cells.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="includeNested">If true, includes tables nested within other table cells</param>
    /// <returns>All table nodes found</returns>
    public static IEnumerable<DocumentNode> FindAllTables(this DocumentNode root, bool includeNested = true)
    {
        foreach (var node in root.FindAll(n => n.Type == ContentType.Table))
        {
            yield return node;

            if (includeNested)
            {
                // Check for nested tables within this table's cells
                var tableData = node.GetTableData();
                if (tableData is not null)
                {
                    foreach (var nestedTable in tableData.FindNestedTables())
                    {
                        yield return nestedTable;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Finds nested tables within a table's cells.
    /// </summary>
    /// <param name="tableData">The table data to search</param>
    /// <returns>All nested table nodes found within cells</returns>
    public static IEnumerable<DocumentNode> FindNestedTables(this TableData tableData)
    {
        foreach (var row in tableData.Rows)
        {
            foreach (var cell in row.Cells)
            {
                foreach (var content in cell.Content)
                {
                    if (content.Type == ContentType.Table)
                    {
                        yield return content;

                        // Recursively find nested tables
                        var nestedData = content.GetTableData();
                        if (nestedData is not null)
                        {
                            foreach (var deepNested in nestedData.FindNestedTables())
                            {
                                yield return deepNested;
                            }
                        }
                    }
                }
            }
        }
    }

    #endregion

    #region Cell access

    /// <summary>
    /// Gets a cell at the specified position.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="row">Zero-based row index</param>
    /// <param name="column">Zero-based column index</param>
    /// <returns>The cell at the position, or null if not found</returns>
    public static TableCell? GetCell(this DocumentNode tableNode, int row, int column)
        => tableNode.GetTableData()?.GetCell(row, column);

    /// <summary>
    /// Gets all cells in a specific row.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="rowIndex">Zero-based row index</param>
    /// <returns>All cells in the row, or empty if row not found</returns>
    public static IEnumerable<TableCell> GetRowCells(this DocumentNode tableNode, int rowIndex)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null || rowIndex < 0 || rowIndex >= tableData.RowCount)
            yield break;

        foreach (var cell in tableData.Rows[rowIndex].Cells)
        {
            yield return cell;
        }
    }

    /// <summary>
    /// Gets all cells in a specific column.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="columnIndex">Zero-based column index</param>
    /// <returns>All cells in the column</returns>
    public static IEnumerable<TableCell> GetColumnCells(this DocumentNode tableNode, int columnIndex)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null)
            yield break;

        for (var row = 0; row < tableData.RowCount; row++)
        {
            var cell = tableData.GetCell(row, columnIndex);
            if (cell is not null)
                yield return cell;
        }
    }

    /// <summary>
    /// Iterates over all cells in the table.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <returns>All cells in the table</returns>
    public static IEnumerable<TableCell> GetAllCells(this DocumentNode tableNode)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null)
            yield break;

        foreach (var row in tableData.Rows)
        {
            foreach (var cell in row.Cells)
            {
                yield return cell;
            }
        }
    }

    /// <summary>
    /// Iterates over all cells with their row and column indices.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <returns>Tuples of (row, column, cell)</returns>
    public static IEnumerable<(int Row, int Column, TableCell Cell)> EnumerateCells(this DocumentNode tableNode)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null)
            yield break;

        foreach (var row in tableData.Rows)
        {
            foreach (var cell in row.Cells)
            {
                yield return (cell.RowIndex, cell.ColumnIndex, cell);
            }
        }
    }

    #endregion

    #region Cell text manipulation

    /// <summary>
    /// Gets the text content of a cell.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="row">Zero-based row index</param>
    /// <param name="column">Zero-based column index</param>
    /// <returns>The cell text, or null if cell not found</returns>
    public static string? GetCellText(this DocumentNode tableNode, int row, int column)
        => tableNode.GetCell(row, column)?.TextContent;

    /// <summary>
    /// Sets the text content of a cell. Creates a new paragraph node if the cell is empty.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="row">Zero-based row index</param>
    /// <param name="column">Zero-based column index</param>
    /// <param name="text">The text to set</param>
    /// <returns>True if successful, false if cell not found</returns>
    public static bool SetCellText(this DocumentNode tableNode, int row, int column, string text)
    {
        var cell = tableNode.GetCell(row, column);
        return cell?.SetText(text) ?? false;
    }

    /// <summary>
    /// Sets the text content of a cell.
    /// </summary>
    /// <param name="cell">The cell to modify</param>
    /// <param name="text">The text to set</param>
    /// <returns>True if successful</returns>
    /// <remarks>
    /// Setting a cell's text to the empty string empties it, in the saved document as well as in
    /// the model.
    /// </remarks>
    public static bool SetText(this TableCell cell, string text)
    {
        if (cell.Content.Count == 0)
        {
            cell.Content.Add(new DocumentNode(ContentType.Paragraph, text));
            cell.MarkContentChanged();
            return true;
        }

        var firstContent = cell.Content[0];
        firstContent.Text = text;
        firstContent.MarkChanged(nameof(DocumentNode.Text));

        // Plain text replaces any formatted runs the paragraph had.
        if (firstContent.Runs.Count > 0)
        {
            firstContent.Runs.Clear();
            firstContent.MarkRunsChanged();
        }

        // Extra paragraphs in the cell are not part of the new value.
        if (cell.Content.Count > 1)
        {
            cell.Content.RemoveRange(1, cell.Content.Count - 1);
            cell.MarkContentChanged();
        }

        return true;
    }

    /// <summary>
    /// Appends text to a cell's content.
    /// </summary>
    /// <param name="cell">The cell to modify</param>
    /// <param name="text">The text to append</param>
    public static void AppendText(this TableCell cell, string text)
    {
        cell.Content.Add(new DocumentNode(ContentType.Paragraph, text));
        cell.MarkContentChanged();
    }

    /// <summary>
    /// Removes the first occurrence of the specified text from a cell's content.
    /// Searches through all paragraph nodes in the cell.
    /// </summary>
    /// <param name="cell">The cell to modify</param>
    /// <param name="textToRemove">The text to remove</param>
    /// <returns>True if text was found and removed, false otherwise</returns>
    public static bool RemoveText(this TableCell cell, string textToRemove)
    {
        foreach (var content in cell.Content)
        {
            if (content.Type is not (ContentType.Paragraph or ContentType.Heading or ContentType.ListItem))
                continue;

            if (!content.Text.Contains(textToRemove, StringComparison.Ordinal))
                continue;

            content.Text = content.Text.Replace(textToRemove, "", StringComparison.Ordinal);
            content.MarkChanged(nameof(DocumentNode.Text));

            if (content.Runs.Count > 0)
            {
                content.Runs.Clear();
                content.MarkRunsChanged();
            }

            return true;
        }

        return false;
    }

    /// <summary>
    /// Clears all content from a cell.
    /// </summary>
    /// <param name="cell">The cell to clear</param>
    /// <remarks>
    /// The clear is recorded so the writer empties the cell in the saved document. Previously the
    /// writer had no way to tell an emptied cell from one that simply parsed to no content nodes,
    /// and left the original text in place — which matters when the caller was redacting.
    /// </remarks>
    public static void ClearContent(this TableCell cell)
    {
        cell.Content.Clear();
        cell.MarkContentChanged();
    }

    #endregion

    #region Cell styling

    /// <summary>
    /// Sets the paragraph style for all content in a cell.
    /// </summary>
    /// <param name="cell">The cell to modify</param>
    /// <param name="styleId">The style ID (e.g., "Heading1", "Normal")</param>
    public static void SetContentStyle(this TableCell cell, string styleId)
    {
        foreach (var content in cell.Content)
        {
            if (content.Type is ContentType.Paragraph or ContentType.Heading or ContentType.ListItem)
            {
                content.ChangeStyle(styleId);
            }
        }
    }

    /// <summary>
    /// Sets the background shading/fill color of a cell.
    /// </summary>
    /// <param name="cell">The cell to modify</param>
    /// <param name="fillColor">Hex color code (e.g., "FFFF00" for yellow, "auto" for no fill)</param>
    public static void SetShading(this TableCell cell, string fillColor)
    {
        cell.Formatting ??= new TableCellFormatting();
        cell.Formatting.ShadingFill = fillColor;
        cell.Formatting.MarkChanged(nameof(TableCellFormatting.ShadingFill));
    }

    /// <summary>
    /// Sets the vertical alignment of cell content.
    /// </summary>
    /// <param name="cell">The cell to modify</param>
    /// <param name="alignment">Alignment value: "top", "center", or "bottom"</param>
    public static void SetVerticalAlignment(this TableCell cell, string alignment)
    {
        cell.Formatting ??= new TableCellFormatting();
        cell.Formatting.VerticalAlignment = alignment;
        cell.Formatting.MarkChanged(nameof(TableCellFormatting.VerticalAlignment));
    }

    /// <summary>
    /// Sets all borders of a cell.
    /// </summary>
    /// <param name="cell">The cell to modify</param>
    /// <param name="style">Border style (e.g., "single", "double", "dotted")</param>
    /// <param name="size">Border size in eighths of a point (e.g., 4 = 0.5pt)</param>
    /// <param name="color">Hex color code (e.g., "000000" for black)</param>
    public static void SetBorders(this TableCell cell, string style = "single", int size = 4, string color = "auto")
    {
        cell.Formatting ??= new TableCellFormatting();
        var border = new BorderFormatting
        {
            Style = style,
            Size = size.ToString(),
            Color = color
        };
        cell.Formatting.TopBorder = border.Clone();
        cell.Formatting.BottomBorder = border.Clone();
        cell.Formatting.LeftBorder = border.Clone();
        cell.Formatting.RightBorder = border.Clone();

        // Borders are only rewritten on an explicit edit, so record this one.
        cell.Formatting.MarkChanged(nameof(TableCellFormatting.TopBorder));
        cell.Formatting.MarkChanged(nameof(TableCellFormatting.BottomBorder));
        cell.Formatting.MarkChanged(nameof(TableCellFormatting.LeftBorder));
        cell.Formatting.MarkChanged(nameof(TableCellFormatting.RightBorder));
    }

    #endregion

    #region Row operations

    /// <summary>
    /// Gets a row from the table.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="rowIndex">Zero-based row index</param>
    /// <returns>The row, or null if not found</returns>
    public static TableRow? GetRow(this DocumentNode tableNode, int rowIndex)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null || rowIndex < 0 || rowIndex >= tableData.RowCount)
            return null;
        return tableData.Rows[rowIndex];
    }

    /// <summary>
    /// Adds a new row to the end of the table.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="cellTexts">Text content for each cell in the new row. If fewer values are provided than columns, remaining cells will be empty.</param>
    /// <returns>The newly created row, or null if the table node is invalid</returns>
    public static TableRow? AddRow(this DocumentNode tableNode, params string[] cellTexts)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null) return null;

        var newRowIndex = tableData.RowCount;
        var colCount = tableData.ColumnCount;

        var newRow = new TableRow
        {
            RowIndex = newRowIndex,
            Cells = []
        };

        for (var col = 0; col < colCount; col++)
        {
            var text = col < cellTexts.Length ? cellTexts[col] : string.Empty;
            var cell = new TableCell
            {
                RowIndex = newRowIndex,
                ColumnIndex = col,
                Content = [new DocumentNode(ContentType.Paragraph, text)]
            };
            newRow.Cells.Add(cell);
        }

        tableData.Rows.Add(newRow);
        tableData.MarkRowsChanged();

        // Modify OriginalXml to include the new row, preserving table style
        ApplyXmlStructuralChange(tableNode, xmlTable =>
        {
            var lastXmlRow = xmlTable.Elements<WP.TableRow>().LastOrDefault();
            if (lastXmlRow != null)
                xmlTable.Append(CloneRowWithTexts(lastXmlRow, cellTexts, colCount));
        });

        return newRow;
    }

    /// <summary>
    /// Inserts a new row at the specified index, shifting existing rows down.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="rowIndex">Zero-based index at which to insert the row. Must be between 0 and RowCount (inclusive).</param>
    /// <param name="cellTexts">Text content for each cell in the new row. If fewer values are provided than columns, remaining cells will be empty.</param>
    /// <returns>The newly created row, or null if the table node is invalid or the index is out of range</returns>
    public static TableRow? InsertRow(this DocumentNode tableNode, int rowIndex, params string[] cellTexts)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null || rowIndex < 0 || rowIndex > tableData.RowCount)
            return null;

        var colCount = tableData.ColumnCount;

        var newRow = new TableRow
        {
            RowIndex = rowIndex,
            Cells = []
        };

        for (var col = 0; col < colCount; col++)
        {
            var text = col < cellTexts.Length ? cellTexts[col] : string.Empty;
            var cell = new TableCell
            {
                RowIndex = rowIndex,
                ColumnIndex = col,
                Content = [new DocumentNode(ContentType.Paragraph, text)]
            };
            newRow.Cells.Add(cell);
        }

        tableData.Rows.Insert(rowIndex, newRow);
        tableData.MarkRowsChanged();

        // Re-index rows after the insertion point
        for (var i = rowIndex + 1; i < tableData.Rows.Count; i++)
        {
            tableData.Rows[i].RowIndex = i;
            foreach (var cell in tableData.Rows[i].Cells)
            {
                cell.RowIndex = i;
            }
        }

        // Modify OriginalXml to insert the new row at the correct position
        ApplyXmlStructuralChange(tableNode, xmlTable =>
        {
            var xmlRows = xmlTable.Elements<WP.TableRow>().ToList();
            var templateRow = rowIndex < xmlRows.Count ? xmlRows[rowIndex] : xmlRows.LastOrDefault();
            if (templateRow != null)
            {
                var newXmlRow = CloneRowWithTexts(templateRow, cellTexts, colCount);
                if (rowIndex < xmlRows.Count)
                    xmlRows[rowIndex].InsertBeforeSelf(newXmlRow);
                else
                    xmlTable.Append(newXmlRow);
            }
        });

        return newRow;
    }

    /// <summary>
    /// Removes a row from the table at the specified index.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="rowIndex">Zero-based row index of the row to remove</param>
    /// <returns>True if the row was removed, false if not found</returns>
    public static bool RemoveRow(this DocumentNode tableNode, int rowIndex)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null || rowIndex < 0 || rowIndex >= tableData.RowCount)
            return false;

        tableData.Rows.RemoveAt(rowIndex);
        tableData.MarkRowsChanged();

        // Re-index remaining rows and their cells
        for (var i = rowIndex; i < tableData.Rows.Count; i++)
        {
            tableData.Rows[i].RowIndex = i;
            foreach (var cell in tableData.Rows[i].Cells)
            {
                cell.RowIndex = i;
            }
        }

        // Modify OriginalXml to remove the row
        ApplyXmlStructuralChange(tableNode, xmlTable =>
        {
            var xmlRows = xmlTable.Elements<WP.TableRow>().ToList();
            if (rowIndex < xmlRows.Count)
                xmlRows[rowIndex].Remove();
        });

        return true;
    }

    /// <summary>
    /// Sets the header flag on a row (headers repeat on page breaks).
    /// </summary>
    /// <param name="row">The row to modify</param>
    /// <param name="isHeader">Whether this row is a header row</param>
    public static void SetAsHeader(this TableRow row, bool isHeader = true)
    {
        row.IsHeader = isHeader;
        row.Formatting ??= new TableRowFormatting();
        row.Formatting.IsHeader = isHeader;
        row.Formatting.MarkChanged(nameof(TableRowFormatting.IsHeader));
    }

    /// <summary>
    /// Sets shading for all cells in a row.
    /// </summary>
    /// <param name="row">The row to modify</param>
    /// <param name="fillColor">Hex color code</param>
    public static void SetRowShading(this TableRow row, string fillColor)
    {
        foreach (var cell in row.Cells)
        {
            cell.SetShading(fillColor);
        }
    }

    #endregion

    #region Column operations

    /// <summary>
    /// Adds a new column to the end of the table.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="cellTexts">Text content for each cell in the new column. If fewer values are provided than rows, remaining cells will be empty.</param>
    /// <returns>True if the column was added, false if the table node is invalid</returns>
    public static bool AddColumn(this DocumentNode tableNode, params string[] cellTexts)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null) return false;

        var newColIndex = tableData.ColumnCount;
        tableData.ColumnCount++;

        for (var row = 0; row < tableData.RowCount; row++)
        {
            var text = row < cellTexts.Length ? cellTexts[row] : string.Empty;
            var cell = new TableCell
            {
                RowIndex = row,
                ColumnIndex = newColIndex,
                Content = [new DocumentNode(ContentType.Paragraph, text)]
            };
            tableData.Rows[row].Cells.Add(cell);
            tableData.Rows[row].MarkCellsChanged();
        }

        tableData.MarkRowsChanged();

        // Modify OriginalXml to add the new column, preserving table style
        ApplyXmlStructuralChange(tableNode, xmlTable =>
        {
            var xmlRows = xmlTable.Elements<WP.TableRow>().ToList();
            for (var row = 0; row < xmlRows.Count; row++)
            {
                var cells = xmlRows[row].Elements<WP.TableCell>().ToList();
                var templateCell = cells.LastOrDefault();
                var text = row < cellTexts.Length ? cellTexts[row] : string.Empty;
                if (templateCell != null)
                    xmlRows[row].Append(CloneCellWithText(templateCell, text));
            }
            AppendGridColumn(xmlTable);
        });

        return true;
    }

    /// <summary>
    /// Inserts a new column at the specified index, shifting existing columns to the right.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="columnIndex">Zero-based index at which to insert the column. Must be between 0 and ColumnCount (inclusive).</param>
    /// <param name="cellTexts">Text content for each cell in the new column. If fewer values are provided than rows, remaining cells will be empty.</param>
    /// <returns>True if the column was inserted, false if the table node is invalid or the index is out of range</returns>
    public static bool InsertColumn(this DocumentNode tableNode, int columnIndex, params string[] cellTexts)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null || columnIndex < 0 || columnIndex > tableData.ColumnCount)
            return false;

        // Shift existing cells at or after the insertion point
        foreach (var row in tableData.Rows)
        {
            foreach (var cell in row.Cells.Where(c => c.ColumnIndex >= columnIndex))
            {
                cell.ColumnIndex++;
            }
        }

        tableData.ColumnCount++;

        // Insert new cells
        for (var row = 0; row < tableData.RowCount; row++)
        {
            var text = row < cellTexts.Length ? cellTexts[row] : string.Empty;
            var cell = new TableCell
            {
                RowIndex = row,
                ColumnIndex = columnIndex,
                Content = [new DocumentNode(ContentType.Paragraph, text)]
            };

            // Insert in sorted position within the row's cell list
            var insertAt = tableData.Rows[row].Cells.FindIndex(c => c.ColumnIndex > columnIndex);
            if (insertAt < 0)
                tableData.Rows[row].Cells.Add(cell);
            else
                tableData.Rows[row].Cells.Insert(insertAt, cell);

            tableData.Rows[row].MarkCellsChanged();
        }

        tableData.MarkRowsChanged();

        // Modify OriginalXml to insert the new column at the correct position
        ApplyXmlStructuralChange(tableNode, xmlTable =>
        {
            var xmlRows = xmlTable.Elements<WP.TableRow>().ToList();
            for (var row = 0; row < xmlRows.Count; row++)
            {
                var cells = xmlRows[row].Elements<WP.TableCell>().ToList();
                var text = row < cellTexts.Length ? cellTexts[row] : string.Empty;
                var templateCell = columnIndex < cells.Count ? cells[columnIndex] : cells.LastOrDefault();
                if (templateCell != null)
                {
                    var newCell = CloneCellWithText(templateCell, text);
                    if (columnIndex < cells.Count)
                        cells[columnIndex].InsertBeforeSelf(newCell);
                    else
                        xmlRows[row].Append(newCell);
                }
            }
            InsertGridColumnAt(xmlTable, columnIndex);
        });

        return true;
    }

    /// <summary>
    /// Removes a column from the table at the specified index.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="columnIndex">Zero-based column index of the column to remove</param>
    /// <returns>True if the column was removed, false if not found</returns>
    /// <remarks>
    /// Column indices are logical: a cell spanning two columns occupies two of them but is a single
    /// cell in the row. Each row is walked by accumulated span to find the cell covering the
    /// requested column, because indexing the row's cells directly picks the wrong one wherever a
    /// horizontal merge sits to the left.
    /// </remarks>
    public static bool RemoveColumn(this DocumentNode tableNode, int columnIndex)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null || columnIndex < 0 || columnIndex >= tableData.ColumnCount)
            return false;

        // Resolve the logical column to a physical cell position per row before mutating anything,
        // so the XML pass and the model pass agree on which cell to remove.
        var physicalIndexPerRow = new List<int>(tableData.Rows.Count);
        foreach (var row in tableData.Rows)
        {
            physicalIndexPerRow.Add(FindPhysicalCellIndex(row, columnIndex));
        }

        for (var rowIndex = 0; rowIndex < tableData.Rows.Count; rowIndex++)
        {
            var row = tableData.Rows[rowIndex];
            var physicalIndex = physicalIndexPerRow[rowIndex];
            if (physicalIndex < 0) continue;

            var cell = row.Cells[physicalIndex];

            if (cell.ColSpan > 1)
            {
                // The column is part of a merge: narrow the merged cell rather than deleting it.
                cell.ColSpan--;
                if (cell.Formatting is not null)
                {
                    cell.Formatting.GridSpan = cell.ColSpan;
                }
            }
            else
            {
                row.Cells.RemoveAt(physicalIndex);
                row.MarkCellsChanged();
            }

            foreach (var following in row.Cells.Where(c => c.ColumnIndex > columnIndex))
            {
                following.ColumnIndex--;
            }
        }

        tableData.ColumnCount--;
        tableData.MarkRowsChanged();

        // Modify OriginalXml to remove the column
        ApplyXmlStructuralChange(tableNode, xmlTable =>
        {
            var xmlRows = xmlTable.Elements<WP.TableRow>().ToList();
            for (var rowIndex = 0; rowIndex < xmlRows.Count && rowIndex < physicalIndexPerRow.Count; rowIndex++)
            {
                var physicalIndex = physicalIndexPerRow[rowIndex];
                if (physicalIndex < 0) continue;

                var cells = xmlRows[rowIndex].Elements<WP.TableCell>().ToList();
                if (physicalIndex >= cells.Count) continue;

                var xmlCell = cells[physicalIndex];
                var gridSpan = xmlCell.TableCellProperties?.GridSpan;
                var span = (int)(gridSpan?.Val?.Value ?? 1);

                if (span > 1)
                {
                    if (span - 1 == 1)
                    {
                        gridSpan!.Remove();
                    }
                    else
                    {
                        gridSpan!.Val = span - 1;
                    }
                }
                else
                {
                    xmlCell.Remove();
                }
            }

            RemoveGridColumnAt(xmlTable, columnIndex);
        });

        return true;
    }

    /// <summary>
    /// Finds the position within a row's cell list of the cell covering a logical column.
    /// </summary>
    /// <returns>The position, or -1 when no cell covers that column.</returns>
    private static int FindPhysicalCellIndex(TableRow row, int columnIndex)
    {
        for (var i = 0; i < row.Cells.Count; i++)
        {
            var cell = row.Cells[i];
            var span = Math.Max(1, cell.ColSpan);
            if (columnIndex >= cell.ColumnIndex && columnIndex < cell.ColumnIndex + span)
            {
                return i;
            }
        }

        return -1;
    }

    #endregion

    #region Table-level operations

    /// <summary>
    /// Gets the dimensions of a table.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <returns>Tuple of (RowCount, ColumnCount)</returns>
    public static (int Rows, int Columns) GetDimensions(this DocumentNode tableNode)
    {
        var tableData = tableNode.GetTableData();
        return tableData is null ? (0, 0) : (tableData.RowCount, tableData.ColumnCount);
    }

    /// <summary>
    /// Sets the table alignment (left, center, right).
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="alignment">Alignment value: "Left", "Center", or "Right"</param>
    /// <remarks>
    /// The change is recorded on the table's formatting and applied to the table's own
    /// <c>w:tblPr</c> on save, not to the <c>w:jc</c> of a paragraph inside one of its cells.
    /// </remarks>
    public static void SetTableAlignment(this DocumentNode tableNode, string alignment)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null) return;

        tableData.Formatting ??= new TableFormatting();
        tableData.Formatting.Alignment = alignment;
        tableData.Formatting.MarkChanged(nameof(TableFormatting.Alignment));
    }

    /// <summary>
    /// Converts table content to a 2D string array.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <returns>2D array of cell text content</returns>
    public static string[,]? ToTextArray(this DocumentNode tableNode)
        => tableNode.GetTableData()?.ToTextArray();

    /// <summary>
    /// Prints a simple text representation of the table.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <returns>Text representation of the table</returns>
    public static string ToTextRepresentation(this DocumentNode tableNode)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null) return "[Empty Table]";

        var lines = new List<string>();
        var colWidths = new int[tableData.ColumnCount];

        // Calculate column widths
        for (var col = 0; col < tableData.ColumnCount; col++)
        {
            for (var row = 0; row < tableData.RowCount; row++)
            {
                var cell = tableData.GetCell(row, col);
                var text = cell?.TextContent ?? "";
                colWidths[col] = Math.Max(colWidths[col], text.Length);
            }
            colWidths[col] = Math.Max(colWidths[col], 3); // Minimum width
        }

        // Build table
        var separator = "+" + string.Join("+", colWidths.Select(w => new string('-', w + 2))) + "+";
        lines.Add(separator);

        for (var row = 0; row < tableData.RowCount; row++)
        {
            var rowText = "|";
            for (var col = 0; col < tableData.ColumnCount; col++)
            {
                var cell = tableData.GetCell(row, col);
                var text = (cell?.TextContent ?? "").Replace("\n", " ");
                if (text.Length > colWidths[col])
                    text = text[..(colWidths[col] - 2)] + "..";
                rowText += " " + text.PadRight(colWidths[col]) + " |";
            }
            lines.Add(rowText);
            lines.Add(separator);
        }

        return string.Join("\n", lines);
    }

    #endregion

    #region Nested table helpers

    /// <summary>
    /// Checks if a cell contains a nested table.
    /// </summary>
    /// <param name="cell">The cell to check</param>
    /// <returns>True if the cell contains at least one table</returns>
    public static bool HasNestedTable(this TableCell cell)
        => cell.Content.Any(c => c.Type == ContentType.Table);

    /// <summary>
    /// Gets all nested tables within a cell.
    /// </summary>
    /// <param name="cell">The cell to search</param>
    /// <returns>All table nodes within the cell</returns>
    public static IEnumerable<DocumentNode> GetNestedTables(this TableCell cell)
        => cell.Content.Where(c => c.Type == ContentType.Table);

    /// <summary>
    /// Gets the first nested table in a cell.
    /// </summary>
    /// <param name="cell">The cell to search</param>
    /// <returns>The first nested table, or null if none found</returns>
    public static DocumentNode? GetFirstNestedTable(this TableCell cell)
        => cell.Content.FirstOrDefault(c => c.Type == ContentType.Table);

    #endregion

    #region XML structural modification helpers

    /// <summary>
    /// Parses the table's OriginalXml, applies a structural modification, and stores the result back.
    /// If OriginalXml is not set, this is a no-op (the writer will build from TableData).
    /// </summary>
    private static void ApplyXmlStructuralChange(DocumentNode tableNode, Action<WP.Table> modifier)
    {
        if (string.IsNullOrEmpty(tableNode.OriginalXml)) return;
        var table = new WP.Table(tableNode.OriginalXml);
        modifier(table);
        tableNode.OriginalXml = table.OuterXml;
    }

    /// <summary>
    /// Clones an XML table row, preserving cell and paragraph formatting but replacing text content.
    /// </summary>
    private static WP.TableRow CloneRowWithTexts(WP.TableRow template, string[] cellTexts, int colCount)
    {
        var newRow = (WP.TableRow)template.CloneNode(true);

        // Remove header repeat flag from cloned row
        var rowProps = newRow.GetFirstChild<WP.TableRowProperties>();
        rowProps?.GetFirstChild<WP.TableHeader>()?.Remove();

        var cells = newRow.Elements<WP.TableCell>().ToList();
        for (var i = 0; i < cells.Count && i < colCount; i++)
        {
            ReplaceCellXmlText(cells[i], i < cellTexts.Length ? cellTexts[i] : string.Empty);
        }

        return newRow;
    }

    /// <summary>
    /// Clones an XML table cell, preserving formatting but replacing text content.
    /// </summary>
    private static WP.TableCell CloneCellWithText(WP.TableCell template, string text)
    {
        var newCell = (WP.TableCell)template.CloneNode(true);
        ReplaceCellXmlText(newCell, text);
        return newCell;
    }

    /// <summary>
    /// Replaces text content of an XML cell while preserving paragraph properties and run formatting.
    /// </summary>
    private static void ReplaceCellXmlText(WP.TableCell xmlCell, string text)
    {
        var paragraphs = xmlCell.Elements<WP.Paragraph>().ToList();

        if (paragraphs.Count > 0)
        {
            var firstPara = paragraphs[0];

            // Get run properties from the first run to preserve font/size/bold/etc.
            var firstRun = firstPara.GetFirstChild<WP.Run>();
            var runProps = firstRun?.RunProperties;

            // Remove all content except ParagraphProperties
            foreach (var child in firstPara.ChildElements.ToList())
            {
                if (child is not WP.ParagraphProperties)
                    child.Remove();
            }

            // Create new run with preserved formatting
            var newRun = new WP.Run(
                new WP.Text(text) { Space = SpaceProcessingModeValues.Preserve }
            );
            if (runProps != null)
                newRun.RunProperties = (WP.RunProperties)runProps.CloneNode(true);
            firstPara.Append(newRun);

            // Remove extra paragraphs
            for (var i = 1; i < paragraphs.Count; i++)
                paragraphs[i].Remove();
        }
        else
        {
            xmlCell.Append(new WP.Paragraph(
                new WP.Run(
                    new WP.Text(text) { Space = SpaceProcessingModeValues.Preserve }
                )
            ));
        }
    }

    /// <summary>
    /// Appends a grid column to the table grid, cloning the width of the last existing column.
    /// </summary>
    private static void AppendGridColumn(WP.Table xmlTable)
    {
        var grid = xmlTable.GetFirstChild<WP.TableGrid>();
        if (grid == null) return;
        var lastGridCol = grid.Elements<WP.GridColumn>().LastOrDefault();
        grid.Append(lastGridCol != null
            ? (WP.GridColumn)lastGridCol.CloneNode(true)
            : new WP.GridColumn());
    }

    /// <summary>
    /// Inserts a grid column at the specified position.
    /// </summary>
    private static void InsertGridColumnAt(WP.Table xmlTable, int columnIndex)
    {
        var grid = xmlTable.GetFirstChild<WP.TableGrid>();
        if (grid == null) return;
        var gridCols = grid.Elements<WP.GridColumn>().ToList();
        if (columnIndex < gridCols.Count)
        {
            gridCols[columnIndex].InsertBeforeSelf(
                (WP.GridColumn)gridCols[columnIndex].CloneNode(true));
        }
        else if (gridCols.Count > 0)
        {
            grid.Append((WP.GridColumn)gridCols.Last().CloneNode(true));
        }
        else
        {
            grid.Append(new WP.GridColumn());
        }
    }

    /// <summary>
    /// Removes a grid column at the specified position.
    /// </summary>
    private static void RemoveGridColumnAt(WP.Table xmlTable, int columnIndex)
    {
        var grid = xmlTable.GetFirstChild<WP.TableGrid>();
        var gridCols = grid?.Elements<WP.GridColumn>().ToList();
        if (gridCols != null && columnIndex < gridCols.Count)
            gridCols[columnIndex].Remove();
    }

    #endregion
}
