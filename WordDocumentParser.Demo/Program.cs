using System;
using System.IO;
using System.Linq;
using WordDocumentParser;
using WordDocumentParser.FormattingModels;

namespace WordDocumentParser.Demo;

/// <summary>
/// Demonstration program showing how to use the Word Document Tree Parser and Writer library.
/// </summary>
class Program
{
    static void Main(string[] args)
    {
        // Example usage with a file path
        string filePath = args.Length > 0 ? args[0] : "C:\\isolated\\content_control_example.docx";

        if (File.Exists(filePath))
        {
            DemoParseAndDisplay(filePath);

            // Demo: Content Controls - Read and Modify
            Console.WriteLine("\n\nContent Control Demo:");
            Console.WriteLine("=====================");
            DemoContentControls(filePath);

            // Demo: Removing Content Controls and Document Properties
            Console.WriteLine("\n\nRemoving Content Controls Demo:");
            Console.WriteLine("================================");
            DemoRemoveContentControls(filePath);

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
    /// Demonstrates reading, modifying, and saving content controls.
    /// </summary>
    static void DemoContentControls(string inputPath)
    {
        using var parser = new WordDocumentTreeParser();
        var documentTree = parser.ParseFromFile(inputPath);

        // 1. Find all nodes with content controls (both block-level and inline)
        Console.WriteLine("1. Finding all content controls...\n");
        var nodesWithControls = documentTree.GetAllContentControls().ToList();
        Console.WriteLine($"   Found {nodesWithControls.Count} node(s) with content controls\n");

        if (nodesWithControls.Count == 0)
        {
            Console.WriteLine("   No content controls found in this document.");
            return;
        }

        // 2. Display details of each content control
        Console.WriteLine("2. Content Control Details:\n");
        var controlIndex = 0;
        foreach (var node in nodesWithControls)
        {
            // Check for block-level content control
            if (node.ContentControlProperties is not null)
            {
                controlIndex++;
                DisplayContentControlDetails(controlIndex, node.ContentControlProperties, "Block-level");
            }

            // Check for inline content controls in runs
            var inlineControls = node.GetInlineContentControlProperties().ToList();
            foreach (var props in inlineControls)
            {
                controlIndex++;
                DisplayContentControlDetails(controlIndex, props, "Inline");
            }
        }

        // 3. Display text with metadata annotations
        Console.WriteLine("3. Text with Metadata (before modification):\n");
        foreach (var node in nodesWithControls)
        {
            Console.WriteLine($"   {node.GetTextWithMetadata()}");
        }
        Console.WriteLine();

        // 4. Modify content control values
        Console.WriteLine("4. Modifying content control values...\n");

        foreach (var node in nodesWithControls)
        {
            // Handle block-level content controls
            if (node.ContentControlProperties is not null)
            {
                ModifyContentControl(node, node.ContentControlProperties);
            }

            // Handle inline content controls
            var inlineControls = node.GetInlineContentControlProperties().ToList();
            foreach (var props in inlineControls)
            {
                ModifyInlineContentControl(node, props);
            }
        }

        // 5. Display text with metadata after modification
        Console.WriteLine("\n5. Text with Metadata (after modification):\n");
        foreach (var node in nodesWithControls)
        {
            Console.WriteLine($"   {node.GetTextWithMetadata()}");
        }
        Console.WriteLine();

        // 6. Save the modified document
        var outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_modified.docx");

        Console.WriteLine("6. Saving modified document...\n");
        documentTree.SaveToFile(outputPath);
        Console.WriteLine($"   Saved to: {outputPath}\n");

        // 7. Verify by re-parsing the saved document
        Console.WriteLine("7. Verifying saved document...\n");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedTree = verifyParser.ParseFromFile(outputPath);

        var verifiedControls = verifiedTree.GetAllContentControls().ToList();
        Console.WriteLine($"   Nodes with content controls in saved document: {verifiedControls.Count}\n");

        Console.WriteLine("   Verified content:");
        foreach (var node in verifiedControls)
        {
            Console.WriteLine($"   {node.GetTextWithMetadata()}");
        }
    }

    /// <summary>
    /// Displays details of a content control.
    /// </summary>
    static void DisplayContentControlDetails(int index, ContentControlProperties props, string location)
    {
        Console.WriteLine($"   [{index}] {location} Content Control:");
        Console.WriteLine($"      Type:  {props.Type}");
        Console.WriteLine($"      ID:    {props.Id}");
        Console.WriteLine($"      Tag:   {props.Tag ?? "(none)"}");
        Console.WriteLine($"      Alias: {props.Alias ?? "(none)"}");
        Console.WriteLine($"      Value: \"{props.Value}\"");

        // Show list items for dropdown/combobox
        if (props.ListItems.Count > 0)
        {
            Console.WriteLine($"      List Items:");
            foreach (var item in props.ListItems)
            {
                Console.WriteLine($"         - \"{item.DisplayText}\" (value: \"{item.Value}\")");
            }
        }

        // Show date info
        if (props.Type == ContentControlType.Date)
        {
            Console.WriteLine($"      Date Format: {props.DateFormat ?? "(none)"}");
            Console.WriteLine($"      Date Value:  {props.DateValue?.ToString("yyyy-MM-dd") ?? "(none)"}");
        }

        // Show checkbox state
        if (props.Type == ContentControlType.Checkbox)
        {
            Console.WriteLine($"      Checked: {props.IsChecked}");
        }

        // Show lock settings
        if (props.LockContentControl || props.LockContents)
        {
            Console.WriteLine($"      Locks: Control={props.LockContentControl}, Contents={props.LockContents}");
        }

        Console.WriteLine();
    }

    /// <summary>
    /// Modifies a block-level content control.
    /// </summary>
    static void ModifyContentControl(DocumentNode node, ContentControlProperties props)
    {
        var oldValue = props.Value;
        var newValue = GetNewValue(props, oldValue);

        // Update the content control value
        node.Text = newValue;
        props.Value = newValue;

        // Update runs if present
        if (node.Runs.Count > 0)
        {
            var formatting = node.Runs.FirstOrDefault()?.Formatting ?? new RunFormatting();
            node.Runs.Clear();
            node.Runs.Add(new FormattedRun(newValue, formatting));
        }

        Console.WriteLine($"   Block control changed: \"{oldValue}\" -> \"{newValue}\"");
    }

    /// <summary>
    /// Modifies an inline content control within a node's runs.
    /// </summary>
    static void ModifyInlineContentControl(DocumentNode node, ContentControlProperties props)
    {
        var oldValue = props.Value;
        var newValue = GetNewValue(props, oldValue);

        // Find and update all runs that belong to this content control
        foreach (var run in node.Runs.Where(r => r.ContentControlProperties == props))
        {
            run.Text = newValue;
        }

        // Update the content control properties
        props.Value = newValue;

        Console.WriteLine($"   Inline control ({props.Type}) changed: \"{oldValue}\" -> \"{newValue}\"");
    }

    /// <summary>
    /// Determines a new value for a content control based on its type.
    /// </summary>
    static string GetNewValue(ContentControlProperties props, string? oldValue)
    {
        switch (props.Type)
        {
            case ContentControlType.DropDownList:
            case ContentControlType.ComboBox:
                // Pick a different item from the list if available
                var otherItem = props.ListItems.FirstOrDefault(i => i.Value != oldValue);
                return otherItem?.DisplayText ?? otherItem?.Value ?? "Selected Item";

            case ContentControlType.Date:
                props.DateValue = DateTime.Now;
                return DateTime.Now.ToString("yyyy-MM-dd");

            case ContentControlType.Checkbox:
                // Toggle the checkbox
                props.IsChecked = !(props.IsChecked ?? false);
                return props.IsChecked == true ? "☒" : "☐";

            case ContentControlType.PlainText:
            case ContentControlType.RichText:
                return $"Modified: {oldValue}";

            default:
                return "Modified Value";
        }
    }

    /// <summary>
    /// Demonstrates removing content controls and document property fields from a document.
    /// The text content is preserved, but the control/field wrapper is removed.
    /// </summary>
    static void DemoRemoveContentControls(string inputPath)
    {
        using var parser = new WordDocumentTreeParser();
        var documentTree = parser.ParseFromFile(inputPath);

        // 1. Display current content controls and document properties
        Console.WriteLine("1. Current content controls and document properties:\n");

        var contentControls = documentTree.GetAllContentControls().ToList();
        Console.WriteLine($"   Content Controls: {contentControls.Count}");
        foreach (var node in contentControls)
        {
            Console.WriteLine($"      - {node.GetTextWithMetadata()}");
        }

        var docPropertyNodes = documentTree.GetNodesWithDocumentPropertyFields().ToList();
        Console.WriteLine($"\n   Document Property Fields: {docPropertyNodes.Count}");
        foreach (var node in docPropertyNodes)
        {
            Console.WriteLine($"      - {node.GetTextWithMetadata()}");
        }

        // 2. Demonstrate removing a specific content control by tag
        Console.WriteLine("\n2. Removing content control with tag 'combobox'...\n");

        var comboBoxRemoved = documentTree.RemoveContentControlByTag("combobox");
        Console.WriteLine($"   Removed: {comboBoxRemoved}");

        // Show the node after removal - text should remain but no content control
        var comboBoxNode = documentTree.FindFirst(n => n.Text.Contains("combobox"));
        if (comboBoxNode != null)
        {
            Console.WriteLine($"   After removal: {comboBoxNode.GetTextWithMetadata()}");
        }

        // 3. Demonstrate removing a content control by alias
        Console.WriteLine("\n3. Removing content control with alias 'dropdown'...\n");

        var dropdownRemoved = documentTree.RemoveContentControlByAlias("dropdown");
        Console.WriteLine($"   Removed: {dropdownRemoved}");

        // 4. Demonstrate removing a content control directly from a node
        Console.WriteLine("\n4. Removing the first remaining content control directly...\n");

        var remainingControls = documentTree.GetAllContentControls().ToList();
        if (remainingControls.Count > 0)
        {
            var nodeToModify = remainingControls[0];
            Console.WriteLine($"   Before: {nodeToModify.GetTextWithMetadata()}");

            // Get the content control ID (from node or from runs)
            int? ccId = nodeToModify.ContentControlProperties?.Id;
            if (ccId is null)
            {
                // Check inline controls
                var inlineProps = nodeToModify.GetInlineContentControlProperties().FirstOrDefault();
                ccId = inlineProps?.Id;
            }

            nodeToModify.RemoveContentControl(ccId);
            Console.WriteLine($"   After:  {nodeToModify.GetTextWithMetadata()}");
        }

        // 5. Demonstrate removing document property fields
        Console.WriteLine("\n5. Removing document property fields...\n");

        if (docPropertyNodes.Count > 0)
        {
            var nodeWithDocProp = docPropertyNodes[0];
            Console.WriteLine($"   Before: {nodeWithDocProp.GetTextWithMetadata()}");

            nodeWithDocProp.RemoveDocumentPropertyField();
            Console.WriteLine($"   After:  {nodeWithDocProp.GetTextWithMetadata()}");
        }
        else
        {
            Console.WriteLine("   No document property fields found in this document.");
        }

        // 6. Show remaining content controls after removals
        Console.WriteLine("\n6. Remaining content controls after targeted removals:\n");

        var finalControls = documentTree.GetAllContentControls().ToList();
        Console.WriteLine($"   Content Controls: {finalControls.Count}");
        foreach (var node in finalControls)
        {
            Console.WriteLine($"      - {node.GetTextWithMetadata()}");
        }

        // 7. Demonstrate removing ALL remaining content controls
        Console.WriteLine("\n7. Removing ALL remaining content controls...\n");

        var totalRemoved = documentTree.RemoveAllContentControls();
        Console.WriteLine($"   Removed {totalRemoved} content control(s)");

        // 8. Save the modified document
        var outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_controls_removed.docx");

        Console.WriteLine($"\n8. Saving document with content controls removed...\n");
        documentTree.SaveToFile(outputPath);
        Console.WriteLine($"   Saved to: {outputPath}");

        // 9. Verify by re-parsing
        Console.WriteLine("\n9. Verifying saved document...\n");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedTree = verifyParser.ParseFromFile(outputPath);

        var verifiedControls = verifiedTree.GetAllContentControls().ToList();
        Console.WriteLine($"   Content controls in saved document: {verifiedControls.Count}");

        // Show the text content is preserved
        Console.WriteLine("\n   Text content preserved (no control wrappers):");
        foreach (var node in verifiedTree.FindAll(n => n.Type == ContentType.Paragraph))
        {
            var text = node.GetText().Trim();
            if (!string.IsNullOrEmpty(text))
            {
                Console.WriteLine($"      - {text}");
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