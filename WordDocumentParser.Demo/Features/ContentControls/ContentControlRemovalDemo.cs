using System;
using System.Linq;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.ContentControls;

/// <summary>
/// Demonstrates removing content controls and document property fields from a document.
/// The text content is preserved, but the control/field wrapper is removed.
/// </summary>
public static class ContentControlRemovalDemo
{
    public static void Run(string inputPath)
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

        var docPropertyNodes = documentTree.Root.GetNodesWithDocumentPropertyFields().ToList();
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
        var outputPath = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(inputPath)!,
            System.IO.Path.GetFileNameWithoutExtension(inputPath) + "_controls_removed.docx");

        Console.WriteLine("\n8. Saving document with content controls removed...\n");
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
}
