using WordDocumentParser;

namespace WordDocumentParser.Demo
{
    /// <summary>
    /// Helper methods for creating tables in demos.
    /// </summary>
    public static class TableHelper
    {
        /// <summary>
        /// Creates a sample table node with demo data.
        /// </summary>
        public static DocumentNode CreateSampleTable()
        {
            var tableData = new TableData { ColumnCount = 3 };

            // Header row
            var headerRow = new TableRow { RowIndex = 0, IsHeader = true };
            headerRow.Cells.Add(CreateCell(0, 0, "Step"));
            headerRow.Cells.Add(CreateCell(0, 1, "Description"));
            headerRow.Cells.Add(CreateCell(0, 2, "Duration"));
            tableData.Rows.Add(headerRow);

            // Data rows
            var row1 = new TableRow { RowIndex = 1 };
            row1.Cells.Add(CreateCell(1, 0, "1"));
            row1.Cells.Add(CreateCell(1, 1, "Data Collection"));
            row1.Cells.Add(CreateCell(1, 2, "2 weeks"));
            tableData.Rows.Add(row1);

            var row2 = new TableRow { RowIndex = 2 };
            row2.Cells.Add(CreateCell(2, 0, "2"));
            row2.Cells.Add(CreateCell(2, 1, "Analysis"));
            row2.Cells.Add(CreateCell(2, 2, "3 weeks"));
            tableData.Rows.Add(row2);

            var row3 = new TableRow { RowIndex = 3 };
            row3.Cells.Add(CreateCell(3, 0, "3"));
            row3.Cells.Add(CreateCell(3, 1, "Report Writing"));
            row3.Cells.Add(CreateCell(3, 2, "1 week"));
            tableData.Rows.Add(row3);

            var tableNode = new DocumentNode(ContentType.Table, $"[Table: {tableData.RowCount}x{tableData.ColumnCount}]");
            tableNode.Metadata["TableData"] = tableData;
            tableNode.Metadata["RowCount"] = tableData.RowCount;
            tableNode.Metadata["ColumnCount"] = tableData.ColumnCount;

            return tableNode;
        }

        private static TableCell CreateCell(int row, int col, string text)
        {
            var cell = new TableCell { RowIndex = row, ColumnIndex = col };
            cell.Content.Add(new DocumentNode(ContentType.Paragraph, text));
            return cell;
        }
    }
}
