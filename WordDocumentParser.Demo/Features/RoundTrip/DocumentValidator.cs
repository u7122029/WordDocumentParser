using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace WordDocumentParser.Demo.Features.RoundTrip;

/// <summary>
/// Provides document validation functionality for demos.
/// </summary>
public static class DocumentValidator
{
    /// <summary>
    /// The Office version demos validate against.
    /// </summary>
    /// <remarks>
    /// Stated once here so the demos and the documentation agree. The validator used to default to
    /// the SDK's own default while the examples alongside it passed Office 2019, which meant the two
    /// could disagree about whether the same document was valid.
    /// </remarks>
    public const FileFormatVersions TargetVersion = FileFormatVersions.Office2019;

    /// <summary>
    /// Validates a Word document and reports any errors to the console.
    /// </summary>
    /// <param name="filePath">The document to validate.</param>
    /// <param name="version">The Office version to validate against.</param>
    /// <returns>True when the document produced no validation errors.</returns>
    public static bool ValidateAndReport(string filePath, FileFormatVersions version = TargetVersion)
    {
        Console.WriteLine($"\nValidating document against {version}...");

        List<string> errors;
        try
        {
            using var doc = WordprocessingDocument.Open(filePath, false);

            // Render each error while the package is still open: ValidationErrorInfo resolves its
            // part and node lazily, so reading them after disposal throws.
            errors = new OpenXmlValidator(version).Validate(doc).Select(Describe).ToList();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Validation could not run: {ex.GetType().Name}: {ex.Message}");
            return false;
        }

        if (errors.Count == 0)
        {
            Console.WriteLine("Document is valid - no errors found.");
            return true;
        }

        Console.WriteLine($"Found {errors.Count} validation errors:");
        foreach (var error in errors.Take(20))
        {
            Console.WriteLine($"  - {error}");
        }

        if (errors.Count > 20)
        {
            Console.WriteLine($"  ... and {errors.Count - 20} more errors");
        }

        return false;
    }

    /// <summary>
    /// Renders one validation error, including a snippet of the offending element.
    /// </summary>
    private static string Describe(ValidationErrorInfo error)
    {
        var location = error.Part?.Uri?.ToString() ?? "package";
        var description = $"[{location}] {error.Description}";

        if (error.Node is null) return description;

        var xml = error.Node.OuterXml;
        if (xml.Length > 100) xml = xml[..100] + "...";
        return $"{description}{Environment.NewLine}    Node: {xml}";
    }
}
