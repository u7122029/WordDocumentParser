using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace WordDocumentParser.Demo.Features.RoundTrip;

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
        try
        {
            using var doc = WordprocessingDocument.Open(filePath, false);
            var validator = new OpenXmlValidator();

            System.Collections.Generic.IEnumerable<ValidationErrorInfo> errors;
            try
            {
                errors = validator.Validate(doc).ToList(); // ToList to materialize and catch errors
            }
            catch (NullReferenceException)
            {
                // OpenXML SDK can throw NullReferenceException when validating certain relationship types
                // This is an SDK bug/limitation, not necessarily an invalid document
                Console.WriteLine("Validation encountered an internal error (NullReferenceException in SDK).");
                Console.WriteLine("This can happen with complex documents but doesn't mean the document is unusable.");
                Console.WriteLine("Try opening the document in Microsoft Word to verify it works correctly.");
                return;
            }

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
        catch (Exception ex)
        {
            Console.WriteLine($"Validation failed with error: {ex.Message}");
            Console.WriteLine("Try opening the document in Microsoft Word to verify it works correctly.");
        }
    }
}
