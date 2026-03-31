using System;
using System.IO;
using System.Linq;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.Tables;

/// <summary>
/// Demonstrates how to structurally modify tables: adding rows, adding columns,
/// appending text to cells, and removing text from cells.
/// </summary>
public static class TableModificationDemo
{
    public static void Run(string inputPath)
    {
        Console.WriteLine("=== Table Modification Demo ===\n");

        // Parse the document
        using var parser = new WordDocumentTreeParser();
        var doc = parser.ParseFromFile(inputPath);

        // var tables = doc.FindAllTables(includeNested: false).ToList();
        // if (tables.Count == 0)
        // {
        //     Console.WriteLine("No tables found in the document.");
        //     return;
        // }

        // var table = tables[0];
        var table = doc.FindAll(node =>
        {
            if (node.Type != ContentType.Table) return false;
            var data = node.GetCellText(0, 0)!.Trim();
            return data == "Acronym";
        }).First();
        var originalFirstCellText = table.GetCellText(0, 0)!.Trim();
        var (originalRows, originalCols) = table.GetDimensions();
        Console.WriteLine($"Working with first table: {originalRows} rows x {originalCols} columns");
        Console.WriteLine("\nOriginal table:");
        Console.WriteLine(table.ToTextRepresentation());

        // --- 1. Add a row ---
        Console.WriteLine("\n--- 1. Adding a new row ---");
        var cellTexts = new string[originalCols];
        for (var i = 0; i < originalCols; i++)
            cellTexts[i] = $"New Row Cell {i + 1}";

        var newRow = table.AddRow(cellTexts);
        if (newRow != null)
        {
            var (rows, cols) = table.GetDimensions();
            Console.WriteLine($"Row added. Table is now {rows} rows x {cols} columns");
        }

        // --- 2. Add a column ---
        Console.WriteLine("\n--- 2. Adding a new column ---");
        var colTexts = new string[table.GetDimensions().Rows];
        colTexts[0] = "New Column Header";
        for (var i = 1; i < colTexts.Length; i++)
            colTexts[i] = $"Col Value {i}";

        var columnAdded = table.AddColumn(colTexts);
        if (columnAdded)
        {
            var (rows, cols) = table.GetDimensions();
            Console.WriteLine($"Column added. Table is now {rows} rows x {cols} columns");
        }

        Console.WriteLine("\nTable after adding row and column:");
        Console.WriteLine(table.ToTextRepresentation());

        // --- 3. Insert a row at a specific index ---
        Console.WriteLine("\n--- 3. Inserting a row at index 1 ---");
        var insertRowTexts = new string[table.GetDimensions().Columns];
        for (var i = 0; i < insertRowTexts.Length; i++)
            insertRowTexts[i] = $"Inserted Cell {i + 1}";

        var insertedRow = table.InsertRow(1, insertRowTexts);
        if (insertedRow != null)
        {
            var (rows, cols) = table.GetDimensions();
            Console.WriteLine($"Row inserted at index 1. Table is now {rows} rows x {cols} columns");
        }

        // --- 4. Insert a column at a specific index ---
        Console.WriteLine("\n--- 4. Inserting a column at index 1 ---");
        var insertColTexts = new string[table.GetDimensions().Rows];
        insertColTexts[0] = "Inserted Col Header";
        for (var i = 1; i < insertColTexts.Length; i++)
            insertColTexts[i] = $"Inserted {i}";

        var columnInserted = table.InsertColumn(1, insertColTexts);
        if (columnInserted)
        {
            var (rows, cols) = table.GetDimensions();
            Console.WriteLine($"Column inserted at index 1. Table is now {rows} rows x {cols} columns");
        }

        Console.WriteLine("\nTable after insertions:");
        Console.WriteLine(table.ToTextRepresentation());

        // --- 5. Append text to a cell ---
        Console.WriteLine("\n--- 5. Appending text to cell [0, 0] ---");
        var cell00 = table.GetCell(0, 0);
        if (cell00 != null)
        {
            Console.WriteLine($"Before append: \"{cell00.TextContent}\"");
            cell00.AppendText(" (appended paragraph)");
            Console.WriteLine($"After append:  \"{cell00.TextContent}\"");
        }

        // --- 6. Remove text from a cell ---
        Console.WriteLine("\n--- 6. Removing text from the newly added row's first cell ---");
        var targetCell = table.GetCell(table.GetDimensions().Rows - 1, 0);
        if (targetCell != null)
        {
            Console.WriteLine($"Before remove: \"{targetCell.TextContent}\"");
            targetCell.RemoveText("Cell ");
            Console.WriteLine($"After remove:  \"{targetCell.TextContent}\"");
        }

        Console.WriteLine("\nFinal table:");
        Console.WriteLine(table.ToTextRepresentation());

        // --- 7. Save and verify ---
        var outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_table_modified.docx");

        Console.WriteLine($"\nSaving modified document to: {outputPath}");
        doc.SaveToFile(outputPath);

        // Verify by re-parsing — find the same table we modified
        Console.WriteLine("\nVerifying saved document...");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedDoc = verifyParser.ParseFromFile(outputPath);
        var verifiedTable = verifiedDoc.FindAll(node =>
        {
            if (node.Type != ContentType.Table) return false;
            var text = node.GetCellText(0, 0);
            return text != null && text.Contains(originalFirstCellText);
        }).First();
        var (verifiedRows, verifiedCols) = verifiedTable.GetDimensions();
        Console.WriteLine($"Verified table: {verifiedRows} rows x {verifiedCols} columns");
        Console.WriteLine(verifiedTable.ToTextRepresentation());

        Console.WriteLine($"\n=== Demo Complete ===");
        Console.WriteLine($"Output file: {outputPath}");
    }
}
