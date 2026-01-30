using System;
using System.IO;
using System.Linq;
using WordDocumentParser.Extensions;

namespace WordDocumentParser.Demo.Features.Styles;

/// <summary>
/// Demonstrates how to find and modify paragraph styles using the library's style extensions.
/// </summary>
public static class ParagraphStyleDemo
{
    public static void Run(string inputPath)
    {
        Console.WriteLine("=== Paragraph Style Modification Demo ===\n");

        // Parse the document
        using var parser = new WordDocumentTreeParser();
        WordDocument doc = parser.ParseFromFile(inputPath);

        // 1. Display current style distribution using GetStyleDistribution()
        Console.WriteLine("1. Current style distribution:");
        var styleStats = doc.GetStyleDistribution();
        foreach (var (style, count) in styleStats.OrderByDescending(x => x.Value).Take(10))
        {
            Console.WriteLine($"   {style}: {count} paragraphs");
        }
        if (styleStats.Count > 10)
            Console.WriteLine($"   ... and {styleStats.Count - 10} more styles");

        // 2. Find paragraphs by style using FindByStyle()
        Console.WriteLine("\n2. Finding paragraphs with specific styles:");

        var heading1Paragraphs = doc.FindByStyle("Heading1").ToList();
        Console.WriteLine($"   Found {heading1Paragraphs.Count} paragraphs with 'Heading1' style");
        foreach (var p in heading1Paragraphs.Take(3))
        {
            Console.WriteLine($"      - \"{Truncate(p.GetText(), 60)}\"");
        }

        var bodyTextParagraphs = doc.FindByStyle("BodyText").ToList();
        Console.WriteLine($"\n   Found {bodyTextParagraphs.Count} paragraphs with 'BodyText' style");
        foreach (var p in bodyTextParagraphs.Take(3))
        {
            Console.WriteLine($"      - \"{Truncate(p.GetText(), 60)}\"");
        }

        // 3. Demonstrate single node style change using node.ChangeStyle()
        Console.WriteLine("\n3. Demonstrating style changes:");

        if (heading1Paragraphs.Count > 0)
        {
            var firstHeading1 = heading1Paragraphs[0];
            Console.WriteLine($"\n   a) Changing first 'Heading1' to 'Heading2' using node.ChangeStyle():");
            Console.WriteLine($"      Text: \"{Truncate(firstHeading1.GetText(), 50)}\"");
            Console.WriteLine($"      Before: {firstHeading1.GetStyle() ?? "(no style)"}");

            // Use the extension method to change the style
            firstHeading1.ChangeStyle("Heading2");

            Console.WriteLine($"      After:  {firstHeading1.GetStyle() ?? "(no style)"}");
        }

        // 4. Demonstrate bulk style change using ChangeStyleBulk()
        if (bodyTextParagraphs.Count > 0)
        {
            // First, let's change just 2 BodyText paragraphs individually
            Console.WriteLine($"\n   b) Changing 2 'BodyText' paragraphs to 'Quote':");
            foreach (var p in bodyTextParagraphs.Take(2))
            {
                Console.WriteLine($"      - \"{Truncate(p.GetText(), 40)}\"");
                p.ChangeStyle("Quote");
            }
        }

        // 5. Demonstrate bulk change using doc.ChangeStyleBulk()
        var captionCount = doc.ChangeStyleBulk("Caption", "Subtitle");
        Console.WriteLine($"\n   c) Bulk changed {captionCount} 'Caption' paragraphs to 'Subtitle' using ChangeStyleBulk()");

        // 6. Show updated style distribution
        Console.WriteLine("\n4. Updated style distribution:");
        var updatedStats = doc.GetStyleDistribution();
        foreach (var (style, count) in updatedStats.OrderByDescending(x => x.Value).Take(10))
        {
            Console.WriteLine($"   {style}: {count} paragraphs");
        }

        // 7. Save the modified document
        var outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_styles_modified.docx");

        Console.WriteLine($"\n5. Saving modified document to: {outputPath}");
        doc.SaveToFile(outputPath);

        // 8. Verify by re-parsing
        Console.WriteLine("\n6. Verifying saved document:");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedDoc = verifyParser.ParseFromFile(outputPath);

        // Use HasStyle() to verify specific changes
        var newHeading2 = verifiedDoc.FindByStyle("Heading2").FirstOrDefault(p => p.GetText().Contains("Introduction"));
        if (newHeading2 != null)
        {
            Console.WriteLine($"   Verified: 'Introduction' now has style '{newHeading2.GetStyle()}'");
        }

        var quoteCount = verifiedDoc.FindByStyle("Quote").Count();
        Console.WriteLine($"   Verified: Document now has {quoteCount} paragraphs with 'Quote' style");

        var subtitleCount = verifiedDoc.FindByStyle("Subtitle").Count();
        Console.WriteLine($"   Verified: Document now has {subtitleCount} paragraphs with 'Subtitle' style");

        Console.WriteLine("\n=== Demo Complete ===");
    }

    private static string Truncate(string text, int maxLength)
    {
        text = text.Replace("\n", " ").Replace("\r", "").Trim();
        return text.Length <= maxLength ? text : text[..(maxLength - 3)] + "...";
    }
}
