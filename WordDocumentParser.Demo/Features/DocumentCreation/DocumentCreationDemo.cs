using System;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.DocumentCreation;

/// <summary>
/// Demonstrates creating a document from scratch using the tree structure.
/// </summary>
public static class DocumentCreationDemo
{
    public static void Run()
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

        // Create the WordDocument wrapper
        var document = new WordDocument(root);

        // Save the document
        var outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "SampleDocument.docx");
        document.SaveToFile(outputPath);
        Console.WriteLine($"Sample document created: {outputPath}");

        // Display the tree structure
        Console.WriteLine("\nDocument Tree Structure:");
        Console.WriteLine(document.ToTreeString());
    }
}
