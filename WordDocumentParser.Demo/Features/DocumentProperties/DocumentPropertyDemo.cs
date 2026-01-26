using System;
using System.IO;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.DocumentProperties;

/// <summary>
/// Demonstrates the document properties API: get, set, and delete properties.
/// </summary>
public static class DocumentPropertyDemo
{
    public static void Run(string inputPath)
    {
        Console.WriteLine("=== Document Properties Demo ===\n");

        // Parse the document
        using var parser = new WordDocumentTreeParser();
        WordDocument doc = parser.ParseFromFile(inputPath);

        // 1. Display existing properties
        Console.WriteLine("1. Existing properties:");
        foreach (var (name, value) in doc.GetAllProperties())
        {
            Console.WriteLine($"   {name} = \"{value}\"");
        }

        // 2. Read properties using indexer
        Console.WriteLine("\n2. Reading properties via indexer:");
        var hasAnyProperties = false;
        foreach (var (name, _) in doc.GetAllProperties())
        {
            hasAnyProperties = true;
            Console.WriteLine($"   doc[\"{name}\"] = \"{doc[name] ?? "(not set)"}\"");
        }

        if (!hasAnyProperties)
        {
            Console.WriteLine("   (no properties found)");
        }

        // 3. Set built-in properties
        Console.WriteLine("\n3. Setting built-in properties:");
        doc["Title"] = "Demo Document";
        doc["Author"] = "WordDocumentParser";
        doc["Subject"] = "Property API Demo";
        doc["Company"] = "Demo Corp";
        Console.WriteLine($"   Title = \"{doc["Title"]}\"");
        Console.WriteLine($"   Author = \"{doc["Author"]}\"");
        Console.WriteLine($"   Subject = \"{doc["Subject"]}\"");
        Console.WriteLine($"   Company = \"{doc["Company"]}\"");

        // 4. Set custom properties (any unknown property name becomes custom)
        Console.WriteLine("\n4. Setting custom properties:");
        doc["ProjectCode"] = "DEMO-001";
        doc["Department"] = "Engineering";
        doc["ReviewStatus"] = "Draft";
        Console.WriteLine($"   ProjectCode = \"{doc["ProjectCode"]}\"");
        Console.WriteLine($"   Department = \"{doc["Department"]}\"");
        Console.WriteLine($"   ReviewStatus = \"{doc["ReviewStatus"]}\"");

        // 5. Access custom properties directly
        Console.WriteLine("\n5. Custom properties dictionary:");
        foreach (var (name, value) in doc.CustomProperties)
        {
            Console.WriteLine($"   {name} = \"{value}\"");
        }

        // 6. Update a custom property
        Console.WriteLine("\n6. Updating custom property:");
        Console.WriteLine($"   Before: ReviewStatus = \"{doc["ReviewStatus"]}\"");
        doc["ReviewStatus"] = "Reviewed";
        Console.WriteLine($"   After:  ReviewStatus = \"{doc["ReviewStatus"]}\"");

        // 7. Delete properties (set to null or use RemoveProperty)
        Console.WriteLine("\n7. Deleting properties:");
        Console.WriteLine($"   HasProperty(\"Department\"): {doc.HasProperty("Department")}");
        doc["Department"] = null; // Delete via indexer
        Console.WriteLine($"   After delete: HasProperty(\"Department\"): {doc.HasProperty("Department")}");

        doc.RemoveProperty("ReviewStatus"); // Delete via method
        Console.WriteLine($"   After RemoveProperty: HasProperty(\"ReviewStatus\"): {doc.HasProperty("ReviewStatus")}");

        // 8. Save and verify
        var outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_properties_demo.docx");

        Console.WriteLine($"\n8. Saving to: {outputPath}");
        doc.SaveToFile(outputPath);

        // 9. Re-parse and verify properties persisted
        Console.WriteLine("\n9. Verifying saved document:");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedDoc = verifyParser.ParseFromFile(outputPath);

        Console.WriteLine("   All properties in saved document:");
        foreach (var (name, value) in verifiedDoc.GetAllProperties())
        {
            Console.WriteLine($"   {name} = \"{value}\"");
        }

        Console.WriteLine("\n=== Demo Complete ===");
    }
}
