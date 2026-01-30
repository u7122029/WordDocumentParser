using System;
using System.IO;
using System.Linq;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.Tables;

/// <summary>
/// Demonstrates how to access and manipulate table data, including nested tables.
/// </summary>
public static class TableParsing
{
    public static void Run(string inputPath)
    {
        Console.WriteLine("=== Table Parsing and Manipulation Demo ===\n");

        // Parse the document
        using var parser = new WordDocumentTreeParser();
        WordDocument doc = parser.ParseFromFile(inputPath);

        // 1. Find all tables in the document (including nested)
        Console.WriteLine("1. Finding all tables in the document:");
        var allTables = doc.FindAllTables(includeNested: true).ToList();
        Console.WriteLine($"   Found {allTables.Count} table(s) total\n");

        // Also show count without nested tables
        var topLevelTables = doc.FindAllTables(includeNested: false).ToList();
        Console.WriteLine($"   Top-level tables: {topLevelTables.Count}");
        Console.WriteLine($"   Nested tables: {allTables.Count - topLevelTables.Count}\n");

        // 2. Display table information for each table
        Console.WriteLine("2. Table details:");
        for (var i = 0; i < allTables.Count; i++)
        {
            var table = allTables[i];
            var (rows, cols) = table.GetDimensions();
            Console.WriteLine($"\n   Table {i + 1}: {rows} rows x {cols} columns");

            // Show text representation
            Console.WriteLine("   " + table.ToTextRepresentation().Replace("\n", "\n   "));
        }

        // 3. Demonstrate cell access
        Console.WriteLine("\n3. Cell access demonstration:");
        if (allTables.Count > 0)
        {
            var firstTable = allTables[0];
            var (rows, cols) = firstTable.GetDimensions();

            // Access specific cell
            if (rows > 0 && cols > 0)
            {
                var cellText = firstTable.GetCellText(0, 0);
                Console.WriteLine($"   Cell [0,0] text: \"{cellText}\"");
            }

            // Get all cells in first row
            Console.WriteLine($"   First row cells:");
            foreach (var cell in firstTable.GetRowCells(0))
            {
                Console.WriteLine($"      Column {cell.ColumnIndex}: \"{Truncate(cell.TextContent, 30)}\"");
            }

            // Get all cells in first column
            if (rows > 1)
            {
                Console.WriteLine($"   First column cells:");
                foreach (var cell in firstTable.GetColumnCells(0))
                {
                    Console.WriteLine($"      Row {cell.RowIndex}: \"{Truncate(cell.TextContent, 30)}\"");
                }
            }
        }

        // 4. Demonstrate cell modification
        Console.WriteLine("\n4. Modifying table cells:");
        if (allTables.Count > 0)
        {
            var table = allTables[0];
            var (rows, cols) = table.GetDimensions();

            // Modify cell text
            if (rows > 0 && cols > 0)
            {
                var originalText = table.GetCellText(0, 0);
                table.SetCellText(0, 0, "[MODIFIED] " + originalText);
                Console.WriteLine($"   Modified cell [0,0]: \"{table.GetCellText(0, 0)}\"");
            }

            // Modify cell styling (apply shading to header row)
            var headerRow = table.GetRow(0);
            if (headerRow is not null)
            {
                Console.WriteLine("   Applying yellow shading to header row...");
                headerRow.SetRowShading("FFFF00"); // Yellow
                headerRow.SetAsHeader(true);
            }

            // Apply different shading to alternating rows
            var tableData = table.GetTableData();
            if (tableData is not null && rows > 1)
            {
                Console.WriteLine("   Applying alternating row shading...");
                for (var r = 1; r < tableData.RowCount; r++)
                {
                    var rowShading = r % 2 == 1 ? "F0F0F0" : "FFFFFF"; // Light gray / white
                    tableData.Rows[r].SetRowShading(rowShading);
                }
            }
        }

        // 5. Demonstrate nested table access
        Console.WriteLine("\n5. Nested table demonstration:");
        var tablesWithNested = allTables.Where(t =>
        {
            var td = t.GetTableData();
            return td is not null && td.Rows.Any(r => r.Cells.Any(c => c.HasNestedTable()));
        }).ToList();

        if (tablesWithNested.Count > 0)
        {
            Console.WriteLine($"   Found {tablesWithNested.Count} table(s) containing nested tables");

            foreach (var parentTable in tablesWithNested)
            {
                var td = parentTable.GetTableData()!;
                foreach (var row in td.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        if (cell.HasNestedTable())
                        {
                            Console.WriteLine($"   Cell [{cell.RowIndex},{cell.ColumnIndex}] contains nested table(s):");

                            foreach (var nestedTable in cell.GetNestedTables())
                            {
                                var (nestedRows, nestedCols) = nestedTable.GetDimensions();
                                Console.WriteLine($"      Nested table: {nestedRows} x {nestedCols}");

                                // Modify nested table
                                Console.WriteLine("      Modifying nested table cell [0,0]...");
                                var originalNestedText = nestedTable.GetCellText(0, 0);
                                if (!string.IsNullOrEmpty(originalNestedText))
                                {
                                    nestedTable.SetCellText(0, 0, "[NESTED-MODIFIED] " + originalNestedText);
                                }
                            }
                        }
                    }
                }
            }
        }
        else
        {
            Console.WriteLine("   No nested tables found in this document.");
        }

        // 6. Demonstrate cell-level formatting
        Console.WriteLine("\n6. Cell formatting demonstration:");
        if (allTables.Count > 0)
        {
            var table = allTables[0];
            var cell = table.GetCell(0, 0);

            if (cell is not null)
            {
                Console.WriteLine("   Setting borders and alignment on cell [0,0]...");
                cell.SetBorders(style: "single", size: 8, color: "000000");
                cell.SetVerticalAlignment("center");
            }
        }

        // 7. Display summary using 2D array
        Console.WriteLine("\n7. Table as 2D array:");
        if (allTables.Count > 0)
        {
            var textArray = allTables[0].ToTextArray();
            if (textArray is not null)
            {
                Console.WriteLine($"   Array dimensions: [{textArray.GetLength(0)}, {textArray.GetLength(1)}]");
                for (var r = 0; r < Math.Min(textArray.GetLength(0), 3); r++)
                {
                    for (var c = 0; c < Math.Min(textArray.GetLength(1), 3); c++)
                    {
                        Console.WriteLine($"   [{r},{c}] = \"{Truncate(textArray[r, c], 30)}\"");
                    }
                }
                if (textArray.GetLength(0) > 3 || textArray.GetLength(1) > 3)
                {
                    Console.WriteLine("   ... (truncated)");
                }
            }
        }

        // 8. Save the modified document
        var outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_tables_modified.docx");

        Console.WriteLine($"\n8. Saving modified document to: {outputPath}");
        doc.SaveToFile(outputPath);

        // 9. Verify by re-parsing
        Console.WriteLine("\n9. Verifying saved document:");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedDoc = verifyParser.ParseFromFile(outputPath);

        var verifiedTables = verifiedDoc.FindAllTables().ToList();
        Console.WriteLine($"   Tables in verified document: {verifiedTables.Count}");

        if (verifiedTables.Count > 0)
        {
            var firstVerifiedCell = verifiedTables[0].GetCellText(0, 0);
            Console.WriteLine($"   First cell text: \"{Truncate(firstVerifiedCell ?? "", 50)}\"");

            // Verify formatting was applied
            var verifiedTableData = verifiedTables[0].GetTableData();
            if (verifiedTableData != null)
            {
                var firstCell = verifiedTableData.GetCell(0, 0);
                if (firstCell?.Formatting != null)
                {
                    Console.WriteLine($"   First cell shading: {firstCell.Formatting.ShadingFill ?? "(none)"}");
                    Console.WriteLine($"   First cell vertical alignment: {firstCell.Formatting.VerticalAlignment ?? "(none)"}");
                }
                else
                {
                    Console.WriteLine("   First cell formatting: (no formatting data)");
                }

                // Check header row
                var headerRow = verifiedTableData.Rows[0];
                Console.WriteLine($"   Header row IsHeader: {headerRow.IsHeader}");

                // Check nested table modifications
                if (firstCell != null && firstCell.HasNestedTable())
                {
                    var nestedTable = firstCell.GetFirstNestedTable();
                    if (nestedTable != null)
                    {
                        var nestedText = nestedTable.GetCellText(0, 0);
                        Console.WriteLine($"   Nested table [0,0] text: \"{Truncate(nestedText ?? "", 50)}\"");
                    }
                }
            }
        }

        Console.WriteLine("\n=== Demo Complete ===");
        Console.WriteLine("\nOpen the modified document in Word to verify:");
        Console.WriteLine("  - First cell should contain '[MODIFIED]' prefix");
        Console.WriteLine("  - Header row should have yellow background (FFFF00)");
        Console.WriteLine("  - Alternating rows should have light gray (F0F0F0) / white background");
        Console.WriteLine($"\nOutput file: {outputPath}");
    }

    private static string Truncate(string text, int maxLength)
    {
        text = text.Replace("\n", " ").Replace("\r", "").Trim();
        return text.Length <= maxLength ? text : text[..(maxLength - 3)] + "...";
    }
}
