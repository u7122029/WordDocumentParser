using System;
using System.IO;
using System.Linq;
using WordDocumentParser.Extensions;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Demo.Features.Fonts;

/// <summary>
/// Demonstrates how to change fonts on runs, text spans, and entire paragraphs.
/// Note: Changing fonts is different from changing paragraph styles - fonts directly
/// control the typeface used to render text.
/// </summary>
public static class FontDemo
{
    public static void Run(string inputPath)
    {
        Console.WriteLine("=== Font Manipulation Demo ===\n");

        // Parse the document
        using var parser = new WordDocumentTreeParser();
        WordDocument doc = parser.ParseFromFile(inputPath);

        // 1. Show current fonts used in the document
        Console.WriteLine("1. Fonts currently used in the document:");
        var fontsUsed = doc.GetAllFontsUsed();
        if (fontsUsed.Count > 0)
        {
            foreach (var font in fontsUsed.OrderBy(f => f))
            {
                Console.WriteLine($"   - {font}");
            }
        }
        else
        {
            Console.WriteLine("   (No explicit fonts set - using default styles)");
        }

        // 2. Change font of a single run
        Console.WriteLine("\n2. Changing font of a single run:");
        var allNodes = doc.Root.FindAll(n =>
            n.Type == Core.ContentType.Paragraph ||
            n.Type == Core.ContentType.Heading ||
            n.Type == Core.ContentType.ListItem).ToList();

        // Also get text content from table cells
        var tableCellParagraphs = new System.Collections.Generic.List<DocumentNode>();
        foreach (var table in doc.FindAllTables())
        {
            foreach (var cell in table.GetAllCells())
            {
                tableCellParagraphs.AddRange(cell.Content.Where(c =>
                    c.Type == Core.ContentType.Paragraph && !string.IsNullOrEmpty(c.Text)));
            }
        }
        allNodes.AddRange(tableCellParagraphs);

        Console.WriteLine($"   Found {allNodes.Count} text-containing node(s)");

        var firstParagraph = allNodes.FirstOrDefault(n => !string.IsNullOrEmpty(n.Text) || !string.IsNullOrEmpty(n.GetText()));
        if (firstParagraph != null)
        {
            // Ensure we have runs to work with
            var nodeText = firstParagraph.GetText();
            if (!firstParagraph.HasFormattedRuns && !string.IsNullOrEmpty(nodeText))
            {
                firstParagraph.Runs.Add(new FormattedRun(nodeText));
            }

            if (firstParagraph.HasFormattedRuns && firstParagraph.Runs.Count > 0)
            {
                var firstRun = firstParagraph.Runs[0];
                Console.WriteLine($"   Original font: {firstRun.GetFont() ?? "(not set)"}");
                Console.WriteLine($"   Text: \"{Truncate(firstRun.Text, 50)}\"");

                // Change to Cascadia Code
                firstRun.SetFont("Cascadia Code");
                Console.WriteLine($"   New font: {firstRun.GetFont()}");
            }
            else
            {
                Console.WriteLine("   (No runs available to modify)");
            }
        }
        else
        {
            Console.WriteLine("   (No paragraphs with text found)");
        }

        // 3. Change font of an entire paragraph
        Console.WriteLine("\n3. Changing font of an entire paragraph:");
        var paragraphs = allNodes.Where(n => !string.IsNullOrEmpty(n.GetText())).ToList();
        if (paragraphs.Count > 1)
        {
            var targetParagraph = paragraphs[1];
            Console.WriteLine($"   Paragraph text: \"{Truncate(targetParagraph.GetText(), 50)}\"");
            Console.WriteLine($"   Fonts before: {string.Join(", ", targetParagraph.GetFontsUsed().DefaultIfEmpty("(none)"))}");

            // Change entire paragraph to Arial
            targetParagraph.SetParagraphFont("Arial");
            Console.WriteLine($"   Fonts after: {string.Join(", ", targetParagraph.GetFontsUsed())}");
        }
        else if (paragraphs.Count == 1)
        {
            var targetParagraph = paragraphs[0];
            Console.WriteLine($"   Paragraph text: \"{Truncate(targetParagraph.GetText(), 50)}\"");
            targetParagraph.SetParagraphFont("Arial");
            Console.WriteLine($"   Changed font to: Arial");
        }
        else
        {
            Console.WriteLine("   (No paragraphs available)");
        }

        // 4. Change font of a specific text span within a paragraph
        Console.WriteLine("\n4. Changing font of a text span (substring):");
        var paragraphWithWords = paragraphs.FirstOrDefault(p => p.GetText().Split(' ', StringSplitOptions.RemoveEmptyEntries).Length >= 2);
        if (paragraphWithWords != null)
        {
            var paragraphText = paragraphWithWords.GetText();
            Console.WriteLine($"   Paragraph: \"{Truncate(paragraphText, 60)}\"");

            // Find a word to change the font of
            var words = paragraphText.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var targetWord = words[0];
            Console.WriteLine($"   Changing font of \"{targetWord}\" to 'Comic Sans MS'...");

            var modified = paragraphWithWords.SetFontForText(targetWord, "Comic Sans MS");
            Console.WriteLine($"   Modified {modified} occurrence(s)");

            // Show the runs after modification
            if (paragraphWithWords.Runs.Count > 0)
            {
                Console.WriteLine("   Runs after modification:");
                foreach (var run in paragraphWithWords.Runs.Take(5))
                {
                    Console.WriteLine($"      \"{Truncate(run.Text, 20)}\" -> {run.GetFont() ?? "(default)"}");
                }
                if (paragraphWithWords.Runs.Count > 5)
                {
                    Console.WriteLine($"      ... and {paragraphWithWords.Runs.Count - 5} more runs");
                }
            }
        }
        else
        {
            Console.WriteLine("   (No paragraph with multiple words found)");
        }

        // 5. Change font by character range
        Console.WriteLine("\n5. Changing font by character range:");
        var paragraphWithLength = paragraphs.FirstOrDefault(p => p.GetText().Length >= 10);
        if (paragraphWithLength != null)
        {
            var paragraphText = paragraphWithLength.GetText();
            Console.WriteLine($"   Paragraph: \"{Truncate(paragraphText, 60)}\"");
            Console.WriteLine($"   Changing characters 0-10 to 'Georgia'...");

            var success = paragraphWithLength.SetFontForRange(0, 10, "Georgia");
            Console.WriteLine($"   Success: {success}");

            if (success && paragraphWithLength.Runs.Count > 0)
            {
                Console.WriteLine("   First few runs:");
                foreach (var run in paragraphWithLength.Runs.Take(3))
                {
                    Console.WriteLine($"      \"{Truncate(run.Text, 25)}\" -> {run.GetFont() ?? "(default)"}");
                }
            }
        }
        else
        {
            Console.WriteLine("   (No paragraph with at least 10 characters found)");
        }

        // 6. Replace one font with another throughout the document
        Console.WriteLine("\n6. Replacing fonts throughout the document:");
        var originalFonts = doc.GetAllFontsUsed().ToList();
        if (originalFonts.Count > 0)
        {
            var fontToReplace = originalFonts[0];
            Console.WriteLine($"   Replacing '{fontToReplace}' with 'Consolas'...");
            var replaced = doc.ReplaceFont(fontToReplace, "Consolas");
            Console.WriteLine($"   Replaced {replaced} run(s)");
        }

        // 7. Set font on all headings
        Console.WriteLine("\n7. Setting font on all headings:");
        var headings = doc.Root.FindAll(n => n.Type == Core.ContentType.Heading).ToList();
        Console.WriteLine($"   Found {headings.Count} heading(s)");
        foreach (var heading in headings)
        {
            heading.SetParagraphFont("Trebuchet MS");
        }
        Console.WriteLine($"   Set all headings to 'Trebuchet MS'");

        // 8. Show updated font usage
        Console.WriteLine("\n8. Fonts now used in the document:");
        var updatedFonts = doc.GetAllFontsUsed();
        foreach (var font in updatedFonts.OrderBy(f => f))
        {
            Console.WriteLine($"   - {font}");
        }

        // 9. Save the modified document
        var outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_fonts_modified.docx");

        Console.WriteLine($"\n9. Saving modified document to: {outputPath}");
        doc.SaveToFile(outputPath);

        // 9.5. Validate the saved document using OpenXML SDK validation
        Console.WriteLine("\n9.5. Validating saved document (OpenXML SDK):");
        using (var validationDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(outputPath, false))
        {
            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            var errors = validator.Validate(validationDoc).ToList();
            if (errors.Count == 0)
            {
                Console.WriteLine("   No validation errors found!");
            }
            else
            {
                Console.WriteLine($"   Found {errors.Count} validation error(s):");
                foreach (var error in errors.Take(30))
                {
                    Console.WriteLine($"   - {error.Description}");
                    Console.WriteLine($"     Part: {error.Part?.Uri}");
                    Console.WriteLine($"     Path: {error.Path?.XPath}");
                }
                if (errors.Count > 30)
                {
                    Console.WriteLine($"   ... and {errors.Count - 30} more errors");
                }
            }
        }

        // 10. Verify by re-parsing
        Console.WriteLine("\n10. Verifying saved document:");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedDoc = verifyParser.ParseFromFile(outputPath);

        var verifiedFonts = verifiedDoc.GetAllFontsUsed();
        Console.WriteLine($"   Fonts in verified document:");
        foreach (var font in verifiedFonts.OrderBy(f => f))
        {
            Console.WriteLine($"      - {font}");
        }

        Console.WriteLine("\n=== Demo Complete ===");
        Console.WriteLine($"\nOutput file: {outputPath}");
        Console.WriteLine("Open the document in Word to see the font changes.");
    }

    private static string Truncate(string text, int maxLength)
    {
        text = text.Replace("\n", " ").Replace("\r", "").Trim();
        return text.Length <= maxLength ? text : text[..(maxLength - 3)] + "...";
    }
}
