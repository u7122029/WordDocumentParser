using System.Text.RegularExpressions;
using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Tables;

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
    public static bool SetText(this TableCell cell, string text)
    {
        if (cell.Content.Count == 0)
        {
            // Create a new paragraph node
            var paraNode = new DocumentNode(ContentType.Paragraph, text);
            cell.Content.Add(paraNode);
        }
        else
        {
            // Update the first paragraph node
            var firstContent = cell.Content[0];
            var oldText = firstContent.Text;
            firstContent.Text = text;

            // Clear runs since we're setting plain text
            firstContent.Runs.Clear();

            // Update OriginalXml if present
            if (!string.IsNullOrEmpty(firstContent.OriginalXml))
            {
                firstContent.OriginalXml = UpdateTextInXml(firstContent.OriginalXml, oldText, text);
            }
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
        var paraNode = new DocumentNode(ContentType.Paragraph, text);
        cell.Content.Add(paraNode);
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

            if (!content.Text.Contains(textToRemove))
                continue;

            var oldText = content.Text;
            content.Text = content.Text.Replace(textToRemove, "");

            // Clear runs since we're modifying plain text
            content.Runs.Clear();

            // Update OriginalXml if present
            if (!string.IsNullOrEmpty(content.OriginalXml))
            {
                content.OriginalXml = UpdateTextInXml(content.OriginalXml, oldText, content.Text);
            }

            return true;
        }

        return false;
    }

    /// <summary>
    /// Clears all content from a cell.
    /// </summary>
    /// <param name="cell">The cell to clear</param>
    public static void ClearContent(this TableCell cell)
    {
        cell.Content.Clear();
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

        // Clear OriginalXml so the writer rebuilds the table from TableData
        tableNode.OriginalXml = null;

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

        // Re-index rows after the insertion point
        for (var i = rowIndex + 1; i < tableData.Rows.Count; i++)
        {
            tableData.Rows[i].RowIndex = i;
            foreach (var cell in tableData.Rows[i].Cells)
            {
                cell.RowIndex = i;
            }
        }

        // Clear OriginalXml so the writer rebuilds the table from TableData
        tableNode.OriginalXml = null;

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

        // Re-index remaining rows and their cells
        for (var i = rowIndex; i < tableData.Rows.Count; i++)
        {
            tableData.Rows[i].RowIndex = i;
            foreach (var cell in tableData.Rows[i].Cells)
            {
                cell.RowIndex = i;
            }
        }

        // Clear OriginalXml so the writer rebuilds the table from TableData
        tableNode.OriginalXml = null;

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
        }

        // Clear OriginalXml so the writer rebuilds the table from TableData
        tableNode.OriginalXml = null;

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
        }

        // Clear OriginalXml so the writer rebuilds the table from TableData
        tableNode.OriginalXml = null;

        return true;
    }

    /// <summary>
    /// Removes a column from the table at the specified index.
    /// </summary>
    /// <param name="tableNode">The table node</param>
    /// <param name="columnIndex">Zero-based column index of the column to remove</param>
    /// <returns>True if the column was removed, false if not found</returns>
    public static bool RemoveColumn(this DocumentNode tableNode, int columnIndex)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null || columnIndex < 0 || columnIndex >= tableData.ColumnCount)
            return false;

        foreach (var row in tableData.Rows)
        {
            row.Cells.RemoveAll(c => c.ColumnIndex == columnIndex);

            // Re-index remaining cells in this row
            foreach (var cell in row.Cells.Where(c => c.ColumnIndex > columnIndex))
            {
                cell.ColumnIndex--;
            }
        }

        tableData.ColumnCount--;

        // Clear OriginalXml so the writer rebuilds the table from TableData
        tableNode.OriginalXml = null;

        return true;
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
    public static void SetTableAlignment(this DocumentNode tableNode, string alignment)
    {
        var tableData = tableNode.GetTableData();
        if (tableData is null) return;

        tableData.Formatting ??= new TableFormatting();
        tableData.Formatting.Alignment = alignment;

        // Also update the OriginalXml if present
        if (!string.IsNullOrEmpty(tableNode.OriginalXml))
        {
            tableNode.OriginalXml = UpdateTableAlignmentInXml(tableNode.OriginalXml, alignment);
        }
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

    #region Private helpers

    private static string UpdateTextInXml(string xml, string oldText, string newText)
    {
        // Try to replace text within <w:t> tags
        var pattern = $@"(<w:t[^>]*>){Regex.Escape(oldText)}(</w:t>)";
        var replacement = $"$1{newText}$2";
        var result = Regex.Replace(xml, pattern, replacement);

        // If no match, try a simpler text replacement (for single <w:t> tags)
        if (result == xml && !string.IsNullOrEmpty(oldText))
        {
            result = xml.Replace($">{oldText}<", $">{newText}<");
        }

        return result;
    }

    private static string UpdateTableAlignmentInXml(string xml, string alignment)
    {
        // Map alignment to Word values
        var wordAlignment = alignment.ToLowerInvariant() switch
        {
            "left" => "left",
            "center" => "center",
            "right" => "right",
            _ => alignment.ToLowerInvariant()
        };

        // Try to replace existing alignment
        var pattern = @"<w:jc\s+w:val=""[^""]*""";
        var replacement = $@"<w:jc w:val=""{wordAlignment}""";

        if (Regex.IsMatch(xml, pattern))
        {
            return Regex.Replace(xml, pattern, replacement);
        }

        // If no existing alignment, try to add one after <w:tblPr>
        var tblPrPattern = @"(<w:tblPr[^>]*>)";
        var match = Regex.Match(xml, tblPrPattern);
        if (match.Success)
        {
            var tblPrTag = match.Groups[1].Value;
            return xml.Replace(tblPrTag, $@"{tblPrTag}<w:jc w:val=""{wordAlignment}""/>");
        }

        return xml;
    }

    #endregion
}
