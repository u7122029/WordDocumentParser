using System;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.RoundTrip;

/// <summary>
/// Demonstrates round-trip parsing and writing of a document.
/// </summary>
public static class RoundTripDemo
{
    public static void Run(string inputPath)
    {
        using var parser = new WordDocumentTreeParser();
        var documentTree = parser.ParseFromFile(inputPath);

        var outputPath = Path.Combine(System.IO.Path.GetDirectoryName(inputPath)!,
        System.IO.Path.GetFileNameWithoutExtension(inputPath) + "_copy.docx");

        documentTree.SaveToFile(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");

        // Validate the saved document
        DocumentValidator.ValidateAndReport(outputPath);

        // Compare specific elements
        Console.WriteLine("\n--- Document Comparison ---");
        DocumentComparison.CompareDocuments(inputPath, outputPath);
    }
}
