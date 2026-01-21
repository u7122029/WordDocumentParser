using System;
using System.IO;
using System.Linq;
using WordDocumentParser;

namespace WordDocumentParser.Demo
{
    /// <summary>
    /// Demonstration program showing how to use the Word Document Tree Parser and Writer library.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage with a file path
            string filePath = args.Length > 0 ? args[0] : "C:\\isolated\\test.docx";

            if (File.Exists(filePath))
            {
                DemoParseAndDisplay(filePath);

                // Demo: Write the parsed document back to a new file
                Console.WriteLine("\n\nWriting Document Demo:");
                Console.WriteLine("======================");
                DemoRoundTrip(filePath);
            }
            else
            {
                Console.WriteLine("Word Document Tree Parser & Writer - Usage Examples");
                Console.WriteLine("===================================================\n");

                // Show example code
                ShowExampleUsage();

                // Demo: Create a document from scratch
                Console.WriteLine("\n\nCreating Sample Document:");
                Console.WriteLine("=========================");
                DemoCreateDocument();
            }
        }

        /// <summary>
        /// Demonstrates parsing a document and displaying its structure.
        /// </summary>
        static void DemoParseAndDisplay(string filePath)
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
                int tableNum = 1;
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

        /// <summary>
        /// Demonstrates round-trip parsing and writing of a document.
        /// </summary>
        static void DemoRoundTrip(string inputPath)
        {
            using var parser = new WordDocumentTreeParser();
            var documentTree = parser.ParseFromFile(inputPath);

            var outputPath = Path.Combine(Path.GetDirectoryName(inputPath)!,
                Path.GetFileNameWithoutExtension(inputPath) + "_copy.docx");

            documentTree.SaveToFile(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");

            // Validate the saved document
            DocumentValidator.ValidateAndReport(outputPath);

            // Compare specific elements
            Console.WriteLine("\n--- Document Comparison ---");
            DocumentComparison.CompareDocuments(inputPath, outputPath);
        }

        /// <summary>
        /// Demonstrates creating a document from scratch using the tree structure.
        /// </summary>
        static void DemoCreateDocument()
        {
            // Create a new document tree programmatically
            var root = new DocumentNode(ContentType.Document, "Sample Document");

            // Add Introduction section
            var intro = new DocumentNode(ContentType.Heading, 1, "Introduction");
            root.AddChild(intro);
            intro.AddChild(new DocumentNode(ContentType.Paragraph,
                "This document was created programmatically using WordDocumentTreeWriter."));
            intro.AddChild(new DocumentNode(ContentType.Paragraph,
                "It demonstrates the ability to generate Word documents from a tree structure."));

            // Add a subsection
            var background = new DocumentNode(ContentType.Heading, 2, "Background");
            intro.AddChild(background);
            background.AddChild(new DocumentNode(ContentType.Paragraph,
                "The document tree structure follows the heading hierarchy of the document."));

            // Add Methods section with a table
            var methods = new DocumentNode(ContentType.Heading, 1, "Methods");
            root.AddChild(methods);
            methods.AddChild(new DocumentNode(ContentType.Paragraph,
                "The following table shows our methodology:"));

            // Create a sample table
            var tableNode = TableHelper.CreateSampleTable();
            methods.AddChild(tableNode);

            // Add Results section with list items
            var results = new DocumentNode(ContentType.Heading, 1, "Results");
            root.AddChild(results);
            results.AddChild(new DocumentNode(ContentType.Paragraph, "Key findings include:"));

            var item1 = new DocumentNode(ContentType.ListItem, "First finding: Improved efficiency");
            item1.Metadata["ListLevel"] = 0;
            item1.Metadata["ListId"] = 1;
            results.AddChild(item1);

            var item2 = new DocumentNode(ContentType.ListItem, "Second finding: Better accuracy");
            item2.Metadata["ListLevel"] = 0;
            item2.Metadata["ListId"] = 1;
            results.AddChild(item2);

            var item3 = new DocumentNode(ContentType.ListItem, "Third finding: Reduced costs");
            item3.Metadata["ListLevel"] = 0;
            item3.Metadata["ListId"] = 1;
            results.AddChild(item3);

            // Add Conclusion
            var conclusion = new DocumentNode(ContentType.Heading, 1, "Conclusion");
            root.AddChild(conclusion);
            conclusion.AddChild(new DocumentNode(ContentType.Paragraph,
                "This demonstrates the round-trip capability of parsing and writing Word documents."));

            // Save the document
            var outputPath = Path.Combine(Environment.CurrentDirectory, "SampleDocument.docx");
            root.SaveToFile(outputPath);
            Console.WriteLine($"Sample document created: {outputPath}");

            // Display the tree structure
            Console.WriteLine("\nDocument Tree Structure:");
            Console.WriteLine(root.ToTreeString());
        }

        static void ShowExampleUsage()
        {
            Console.WriteLine(@"
// Basic Usage - Parse a Word document
using var parser = new WordDocumentTreeParser();
var documentTree = parser.ParseFromFile(""document.docx"");

// The tree structure follows heading hierarchy:
// Document (root)
//   +-- H1: Introduction
//   |     +-- Paragraph: Some text...
//   |     +-- H2: Background
//   |     |     +-- Paragraph: More text...
//   |     |     +-- Table: [Table: 3x4]
//   |     +-- H2: Purpose
//   |           +-- Paragraph: Purpose text...
//   +-- H1: Methods
//         +-- H2: Data Collection
//         |     +-- Image: [Image: figure1.png]
//         +-- H2: Analysis
//               +-- Paragraph: Analysis details...

// Print the tree structure
Console.WriteLine(documentTree.ToTreeString());

// Find all headings
var allHeadings = documentTree.GetAllHeadings();
foreach (var heading in allHeadings)
{
    Console.WriteLine($""H{heading.HeadingLevel}: {heading.Text}"");
}

// Get specific heading level
var h2Headings = documentTree.GetHeadingsAtLevel(2);

// Find a section by name
var methodsSection = documentTree.GetSection(""Methods"");
if (methodsSection != null)
{
    // Get all text under this section
    var sectionText = methodsSection.GetAllText();
    Console.WriteLine($""Methods section content:\n{sectionText}"");
}

// Work with tables
var tables = documentTree.GetAllTables();
foreach (var tableNode in tables)
{
    var tableData = tableNode.GetTableData();
    if (tableData != null)
    {
        // Access as 2D array
        var array = tableData.ToTextArray();
        Console.WriteLine($""Cell [0,0]: {array[0, 0]}"");
    }
}

// Work with images
var images = documentTree.GetAllImages();
foreach (var imageNode in images)
{
    var imageData = imageNode.GetImageData();
    if (imageData?.Data != null)
    {
        // Save image to file
        File.WriteAllBytes($""{imageData.Name}"", imageData.Data);
    }
}

// Navigation
var node = documentTree.FindFirst(n => n.Text.Contains(""specific text""));
if (node != null)
{
    // Get breadcrumb path
    var path = node.GetHeadingPath();

    // Navigate to parent
    var parent = node.Parent;

    // Get siblings
    var siblings = node.GetSiblings();
}

// Statistics
var stats = documentTree.CountByType();
Console.WriteLine($""Total paragraphs: {stats[ContentType.Paragraph]}"");

// ============================================
// WRITING DOCUMENTS
// ============================================

// Save a parsed document back to a new file
documentTree.SaveToFile(""output.docx"");

// Or save to a stream
using var stream = new MemoryStream();
documentTree.SaveToStream(stream);

// Or get as byte array
var bytes = documentTree.ToDocxBytes();

// Create a new document from scratch
var newDoc = new DocumentNode(ContentType.Document, ""My Document"");
var heading = new DocumentNode(ContentType.Heading, 1, ""Chapter 1"");
newDoc.AddChild(heading);
heading.AddChild(new DocumentNode(ContentType.Paragraph, ""This is the first paragraph.""));

// Save the new document
newDoc.SaveToFile(""new_document.docx"");
");
        }
    }
}
