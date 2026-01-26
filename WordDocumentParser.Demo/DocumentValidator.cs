using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace WordDocumentParser.Demo;

/// <summary>
/// Provides document validation functionality for demos.
/// </summary>
public static class DocumentValidator
{
    /// <summary>
    /// Validates a Word document and reports any errors to the console.
    /// </summary>
    public static void ValidateAndReport(string filePath)
    {
        Console.WriteLine("\nValidating document...");
        using var doc = WordprocessingDocument.Open(filePath, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc);

        if (!errors.Any())
        {
            Console.WriteLine("Document is valid - no errors found.");
        }
        else
        {
            Console.WriteLine($"Found {errors.Count()} validation errors:");
            foreach (var error in errors.Take(20))
            {
                Console.WriteLine($"  - {error.Description}");
                if (error.Node != null)
                {
                    var xml = error.Node.OuterXml;
                    if (xml.Length > 100) xml = xml.Substring(0, 100) + "...";
                    Console.WriteLine($"    Node: {xml}");
                }
            }
            if (errors.Count() > 20)
            {
                Console.WriteLine($"  ... and {errors.Count() - 20} more errors");
            }
        }
    }
}