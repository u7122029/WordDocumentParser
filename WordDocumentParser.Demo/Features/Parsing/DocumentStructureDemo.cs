using System;
using System.Linq;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.Parsing;

/// <summary>
/// Demonstrates parsing a document and displaying its structure.
/// </summary>
public static class DocumentStructureDemo
{
    public static void Run(string filePath)
    {
        Console.WriteLine($"Parsing: {filePath}\n");

        using var parser = new WordDocumentTreeParser();
        var documentTree = parser.ParseFromFile(filePath);

        // Display the tree structure
        Console.WriteLine("Document Tree Structure:");
        Console.WriteLine("========================");
        Console.WriteLine(documentTree.ToTreeString());

        // Display statistics
        Console.WriteLine("\nDocument Statistics:");
        Console.WriteLine("====================");
        var counts = documentTree.CountByType();
        foreach (var kvp in counts.OrderBy(k => k.Key.ToString()))
        {
            Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
        }

        // Display table of contents
        var toc = documentTree.GetTableOfContents();
        if (toc.Any())
        {
            Console.WriteLine("\nTable of Contents:");
            Console.WriteLine("==================");
            foreach (var (level, title, _) in toc)
            {
                var indent = new string(' ', (level - 1) * 2);
                Console.WriteLine($"{indent}{level}. {title}");
            }
        }

        // Display tables info
        var tables = documentTree.GetAllTables().ToList();
        if (tables.Any())
        {
            Console.WriteLine($"\nTables Found: {tables.Count}");
            Console.WriteLine("=============");
            var tableNum = 1;
            foreach (var table in tables)
            {
                var tableData = table.GetTableData();
                if (tableData != null)
                {
                    Console.WriteLine($"  Table {tableNum}: {tableData.RowCount} rows x {tableData.ColumnCount} columns");
                    Console.WriteLine($"    Location: {table.GetHeadingPath()}");
                }
                tableNum++;
            }
        }

        // Display images info
        var images = documentTree.GetAllImages().ToList();
        if (images.Any())
        {
            Console.WriteLine($"\nImages Found: {images.Count}");
            Console.WriteLine("=============");
            foreach (var image in images)
            {
                var imageData = image.GetImageData();
                if (imageData != null)
                {
                    Console.WriteLine($"  - {imageData.Name}: {imageData.WidthInches:F1}\" x {imageData.HeightInches:F1}\" ({imageData.ContentType})");
                    Console.WriteLine($"    Location: {image.GetHeadingPath()}");
                }
            }
        }
    }
}
