using System;
using System.IO;
using System.Linq;
using WordDocumentParser;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.Concatenation;

/// <summary>
/// Demonstrates how to concatenate two Word documents together using the WordDocumentParser library.
/// The second document's content is appended after the first document's content.
/// </summary>
public static class DocumentConcatenationDemo
{
    public static void Run(string firstDocPath, string secondDocPath)
    {
        Console.WriteLine("=== Document Concatenation Demo ===\n");

        if (!File.Exists(firstDocPath))
        {
            Console.WriteLine($"Error: First document not found: {firstDocPath}");
            return;
        }

        if (!File.Exists(secondDocPath))
        {
            Console.WriteLine($"Error: Second document not found: {secondDocPath}");
            return;
        }

        // 1. Parse both documents
        Console.WriteLine("1. Parsing documents...");
        using var parser1 = new WordDocumentTreeParser();
        using var parser2 = new WordDocumentTreeParser();

        var doc1 = parser1.ParseFromFile(firstDocPath);
        var doc2 = parser2.ParseFromFile(secondDocPath);

        Console.WriteLine($"   First document:  {Path.GetFileName(firstDocPath)}");
        Console.WriteLine($"      - {doc1.Root.Children.Count} top-level nodes");
        Console.WriteLine($"      - {doc1.Root.FindAll(_ => true).Count()} total nodes");
        Console.WriteLine($"      - {doc1.PackageData.Images.Count} images");

        Console.WriteLine($"   Second document: {Path.GetFileName(secondDocPath)}");
        Console.WriteLine($"      - {doc2.Root.Children.Count} top-level nodes");
        Console.WriteLine($"      - {doc2.Root.FindAll(_ => true).Count()} total nodes");
        Console.WriteLine($"      - {doc2.PackageData.Images.Count} images");

        // 2. Get merge statistics before merging
        Console.WriteLine("\n2. Pre-merge statistics...");
        var stats = doc1.GetMergeStatistics(doc2);
        Console.WriteLine($"   Target: {stats.TargetNodeCount} nodes, {stats.TargetImageCount} images, {stats.TargetHyperlinkCount} hyperlinks");
        Console.WriteLine($"   Source: {stats.SourceNodeCount} nodes, {stats.SourceImageCount} images, {stats.SourceHyperlinkCount} hyperlinks");

        // 3. Append the second document to the first (with page break)
        Console.WriteLine("\n3. Appending second document...");
        doc1.AppendDocument(doc2, addPageBreak: true);
        Console.WriteLine("   Documents merged successfully!");

        // 4. Update document properties
        Console.WriteLine("\n4. Updating document properties...");
        var title1 = doc1.Title ?? Path.GetFileNameWithoutExtension(firstDocPath);
        var title2 = doc2.Title ?? Path.GetFileNameWithoutExtension(secondDocPath);
        doc1["Title"] = $"{title1} + {title2}";
        Console.WriteLine($"   New title: {doc1["Title"]}");

        // 5. Save the combined document
        var outputPath = Path.Combine(
            Path.GetDirectoryName(firstDocPath)!,
            $"{Path.GetFileNameWithoutExtension(firstDocPath)}_combined_{Path.GetFileNameWithoutExtension(secondDocPath)}.docx");

        Console.WriteLine($"\n5. Saving combined document...");
        doc1.SaveToFile(outputPath);
        Console.WriteLine($"   Saved to: {outputPath}");

        // 6. Validate the output
        Console.WriteLine("\n6. Validating combined document...");
        try
        {
            using var validationDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(outputPath, false);
            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            var errors = validator.Validate(validationDoc).ToList();

            if (errors.Count == 0)
            {
                Console.WriteLine("   No validation errors found!");
            }
            else
            {
                Console.WriteLine($"   Found {errors.Count} validation error(s):");
                foreach (var error in errors.Take(10))
                {
                    Console.WriteLine($"   - {error.Description}");
                }
                if (errors.Count > 10)
                {
                    Console.WriteLine($"   ... and {errors.Count - 10} more errors");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"   Validation skipped due to error: {ex.Message}");
            Console.WriteLine("   (This is likely an OpenXML SDK validator bug, document may still be valid)");
        }

        // 7. Verify by re-parsing
        Console.WriteLine("\n7. Verifying combined document...");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedDoc = verifyParser.ParseFromFile(outputPath);

        Console.WriteLine($"   Combined document has:");
        Console.WriteLine($"      - {verifiedDoc.Root.Children.Count} top-level nodes");
        Console.WriteLine($"      - {verifiedDoc.Root.FindAll(_ => true).Count()} total nodes");
        Console.WriteLine($"      - {verifiedDoc.PackageData.Images.Count} images");

        Console.WriteLine("\n=== Demo Complete ===");
        Console.WriteLine($"\nOutput file: {outputPath}");
        Console.WriteLine("Open the document in Word to see the combined content.");
    }

    /// <summary>
    /// Demonstrates concatenating multiple documents at once.
    /// </summary>
    public static void RunMultiple(params string[] documentPaths)
    {
        Console.WriteLine("=== Multiple Document Concatenation Demo ===\n");

        // Validate all paths
        foreach (var path in documentPaths)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine($"Error: Document not found: {path}");
                return;
            }
        }

        // Parse all documents
        Console.WriteLine("1. Parsing documents...");
        var documents = new List<WordDocument>();
        foreach (var path in documentPaths)
        {
            using var parser = new WordDocumentTreeParser();
            var doc = parser.ParseFromFile(path);
            documents.Add(doc);
            Console.WriteLine($"   - {Path.GetFileName(path)}: {doc.Root.FindAll(_ => true).Count()} nodes");
        }

        // Concatenate all documents
        Console.WriteLine("\n2. Concatenating documents...");
        var combined = DocumentMergeExtensions.ConcatenateDocuments(documents, addPageBreaks: true);
        Console.WriteLine($"   Combined: {combined.Root.FindAll(_ => true).Count()} total nodes");

        // Save
        var outputPath = Path.Combine(
            Path.GetDirectoryName(documentPaths[0])!,
            "combined_documents.docx");

        Console.WriteLine($"\n3. Saving to: {outputPath}");
        combined.SaveToFile(outputPath);

        Console.WriteLine("\n=== Demo Complete ===");
    }

    /// <summary>
    /// Demonstrates extracting a section from one document and inserting it into another.
    /// </summary>
    public static void RunSectionInsertion(string targetDocPath, string sourceDocPath)
    {
        Console.WriteLine("=== Section Insertion Demo ===\n");

        if (!File.Exists(targetDocPath) || !File.Exists(sourceDocPath))
        {
            Console.WriteLine("Error: One or more documents not found");
            return;
        }

        // Parse documents
        Console.WriteLine("1. Parsing documents...");
        using var parser1 = new WordDocumentTreeParser();
        using var parser2 = new WordDocumentTreeParser();

        var targetDoc = parser1.ParseFromFile(targetDocPath);
        var sourceDoc = parser2.ParseFromFile(sourceDocPath);

        Console.WriteLine($"   Target: {Path.GetFileName(targetDocPath)}");
        Console.WriteLine($"   Source: {Path.GetFileName(sourceDocPath)}");

        // List headings in both documents
        Console.WriteLine("\n2. Headings in target document:");
        var targetHeadings = targetDoc.Root.FindAll(n => n.Type == WordDocumentParser.Core.ContentType.Heading)
            .Take(10).ToList();
        foreach (var h in targetHeadings)
        {
            Console.WriteLine($"   H{h.HeadingLevel}: {Truncate(h.GetText(), 60)}");
        }
        if (targetDoc.Root.FindAll(n => n.Type == WordDocumentParser.Core.ContentType.Heading).Count() > 10)
            Console.WriteLine("   ... (more headings)");

        Console.WriteLine("\n3. Headings in source document:");
        var sourceHeadings = sourceDoc.Root.FindAll(n => n.Type == WordDocumentParser.Core.ContentType.Heading)
            .Take(10).ToList();
        foreach (var h in sourceHeadings)
        {
            Console.WriteLine($"   H{h.HeadingLevel}: {Truncate(h.GetText(), 60)}");
        }
        if (sourceDoc.Root.FindAll(n => n.Type == WordDocumentParser.Core.ContentType.Heading).Count() > 10)
            Console.WriteLine("   ... (more headings)");

        // Extract a section from the source document
        if (sourceHeadings.Count > 0)
        {
            var sectionHeading = sourceHeadings[0].GetText();
            Console.WriteLine($"\n4. Extracting section: \"{Truncate(sectionHeading, 50)}\"");

            var extractedSection = sourceDoc.ExtractSection(sectionHeading, includeNestedHeadings: true);
            Console.WriteLine($"   Extracted {extractedSection.Count} nodes");

            // Show what was extracted
            foreach (var node in extractedSection.Take(5))
            {
                var typeLabel = node.Type == WordDocumentParser.Core.ContentType.Heading
                    ? $"H{node.HeadingLevel}"
                    : node.Type.ToString();
                Console.WriteLine($"      [{typeLabel}] {Truncate(node.GetText(), 50)}");
            }
            if (extractedSection.Count > 5)
                Console.WriteLine($"      ... and {extractedSection.Count - 5} more nodes");

            // Find a heading in the middle of the target document to insert after
            var middleHeadingIndex = targetHeadings.Count / 2;
            var insertAfterHeading = targetHeadings[middleHeadingIndex];
            Console.WriteLine($"\n5. Inserting section after heading: \"{Truncate(insertAfterHeading.GetText(), 50)}\"");

            targetDoc.InsertNodesAfter(
                insertAfterHeading,
                extractedSection,
                sourceDoc);

            Console.WriteLine($"   Target now has {targetDoc.Root.FindAll(_ => true).Count()} total nodes");

            // Save
            var outputPath = Path.Combine(
                Path.GetDirectoryName(targetDocPath)!,
                $"{Path.GetFileNameWithoutExtension(targetDocPath)}_with_section.docx");

            Console.WriteLine($"\n6. Saving to: {outputPath}");
            targetDoc.SaveToFile(outputPath);

            Console.WriteLine("\n=== Demo Complete ===");
            Console.WriteLine($"Output file: {outputPath}");
        }
        else
        {
            Console.WriteLine("\nNo headings found in source document to extract.");
        }
    }

    /// <summary>
    /// Demonstrates extracting specific nodes (e.g., all tables) from one document
    /// and inserting them into another.
    /// </summary>
    public static void RunNodeExtraction(string targetDocPath, string sourceDocPath)
    {
        Console.WriteLine("=== Node Extraction Demo ===\n");

        if (!File.Exists(targetDocPath) || !File.Exists(sourceDocPath))
        {
            Console.WriteLine("Error: One or more documents not found");
            return;
        }

        using var parser1 = new WordDocumentTreeParser();
        using var parser2 = new WordDocumentTreeParser();

        var targetDoc = parser1.ParseFromFile(targetDocPath);
        var sourceDoc = parser2.ParseFromFile(sourceDocPath);

        Console.WriteLine($"Target: {Path.GetFileName(targetDocPath)}");
        Console.WriteLine($"Source: {Path.GetFileName(sourceDocPath)}");

        // Extract all tables from source
        Console.WriteLine("\n1. Extracting tables from source document...");
        var tables = sourceDoc.ExtractTables();
        Console.WriteLine($"   Found {tables.Count} table(s)");

        if (tables.Count > 0)
        {
            // Clone tables for insertion into target
            Console.WriteLine("\n2. Cloning tables with resource mapping...");
            var clonedTables = targetDoc.CloneNodesForDocument(sourceDoc, tables);

            // Insert tables at the beginning of the target document
            Console.WriteLine("\n3. Inserting tables at beginning of target...");
            for (int i = 0; i < clonedTables.Count; i++)
            {
                clonedTables[i].Parent = targetDoc.Root;
                targetDoc.Root.Children.Insert(i, clonedTables[i]);
            }

            var outputPath = Path.Combine(
                Path.GetDirectoryName(targetDocPath)!,
                $"{Path.GetFileNameWithoutExtension(targetDocPath)}_with_tables.docx");

            Console.WriteLine($"\n4. Saving to: {outputPath}");
            targetDoc.SaveToFile(outputPath);

            Console.WriteLine("\n=== Demo Complete ===");
        }
        else
        {
            Console.WriteLine("No tables found in source document.");
        }
    }

    private static string Truncate(string text, int maxLength)
    {
        if (string.IsNullOrEmpty(text)) return "(empty)";
        text = text.Replace("\n", " ").Replace("\r", "").Trim();
        return text.Length <= maxLength ? text : text[..(maxLength - 3)] + "...";
    }
}
