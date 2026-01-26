using System;

namespace WordDocumentParser.Demo.Features.Examples;

/// <summary>
/// Prints example usage for the demo app.
/// </summary>
public static class ExampleUsageDemo
{
    public static void Show()
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
