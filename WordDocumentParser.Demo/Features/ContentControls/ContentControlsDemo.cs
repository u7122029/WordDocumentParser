using System;
using System.Linq;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using WordDocumentParser.Models.ContentControls;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Demo.Features.ContentControls;

/// <summary>
/// Demonstrates reading, modifying, and saving content controls.
/// </summary>
public static class ContentControlsDemo
{
    public static void Run(string inputPath)
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
        var outputPath = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(inputPath)!,
            System.IO.Path.GetFileNameWithoutExtension(inputPath) + "_modified.docx");

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

    private static void DisplayContentControlDetails(int index, ContentControlProperties props, string location)
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
            Console.WriteLine("      List Items:");
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

    private static void ModifyContentControl(DocumentNode node, ContentControlProperties props)
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

    private static void ModifyInlineContentControl(DocumentNode node, ContentControlProperties props)
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

    private static string GetNewValue(ContentControlProperties props, string? oldValue)
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
                return props.IsChecked == true ? "[X]" : "[ ]";

            case ContentControlType.PlainText:
            case ContentControlType.RichText:
                return $"Modified: {oldValue}";

            default:
                return "Modified Value";
        }
    }
}
