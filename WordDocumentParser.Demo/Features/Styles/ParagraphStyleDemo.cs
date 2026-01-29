using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using WordDocumentParser.Extensions;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Demo.Features.Styles;

/// <summary>
/// Demonstrates how to find and modify paragraph styles in a parsed document.
/// </summary>
public static class ParagraphStyleDemo
{
    public static void Run(string inputPath)
    {
        Console.WriteLine("=== Paragraph Style Modification Demo ===\n");

        // Parse the document
        using var parser = new WordDocumentTreeParser();
        WordDocument doc = parser.ParseFromFile(inputPath);

        // 1. Display current style distribution
        Console.WriteLine("1. Current style distribution:");
        var styleStats = GetStyleDistribution(doc.Root);
        foreach (var (style, count) in styleStats.OrderByDescending(x => x.Value))
        {
            Console.WriteLine($"   {style}: {count} paragraphs");
        }

        // 2. Find paragraphs by style
        Console.WriteLine("\n2. Finding paragraphs with specific styles:");

        var heading1Paragraphs = FindNodesByStyle(doc.Root, "Heading1").ToList();
        Console.WriteLine($"   Found {heading1Paragraphs.Count} paragraphs with 'Heading1' style");
        foreach (var p in heading1Paragraphs.Take(3))
        {
            Console.WriteLine($"      - \"{Truncate(p.GetText(), 60)}\"");
        }

        var bodyTextParagraphs = FindNodesByStyle(doc.Root, "BodyText").ToList();
        Console.WriteLine($"\n   Found {bodyTextParagraphs.Count} paragraphs with 'BodyText' style");
        foreach (var p in bodyTextParagraphs.Take(3))
        {
            Console.WriteLine($"      - \"{Truncate(p.GetText(), 60)}\"");
        }

        // 3. Demonstrate style changes
        Console.WriteLine("\n3. Demonstrating style changes:");

        // Example: Change first Heading1 to Heading2
        if (heading1Paragraphs.Count > 0)
        {
            var firstHeading1 = heading1Paragraphs[0];
            Console.WriteLine($"\n   a) Changing first 'Heading1' to 'Heading2':");
            Console.WriteLine($"      Text: \"{Truncate(firstHeading1.GetText(), 50)}\"");
            Console.WriteLine($"      Before: StyleId = {firstHeading1.ParagraphFormatting?.StyleId ?? "(null)"}");

            ChangeNodeStyle(firstHeading1, "Heading2");

            Console.WriteLine($"      After:  StyleId = {firstHeading1.ParagraphFormatting?.StyleId ?? "(null)"}");
        }

        // Example: Change some BodyText paragraphs to Quote
        var bodyTextToChange = bodyTextParagraphs.Take(2).ToList();
        if (bodyTextToChange.Count > 0)
        {
            Console.WriteLine($"\n   b) Changing {bodyTextToChange.Count} 'BodyText' paragraphs to 'Quote':");
            foreach (var p in bodyTextToChange)
            {
                Console.WriteLine($"      - \"{Truncate(p.GetText(), 40)}\"");
                ChangeNodeStyle(p, "Quote");
            }
        }

        // Example: Bulk change - change all Caption to Subtitle
        var captionCount = ChangeStyleBulk(doc.Root, "Caption", "Subtitle");
        Console.WriteLine($"\n   c) Bulk changed {captionCount} 'Caption' paragraphs to 'Subtitle'");

        // 4. Show updated style distribution
        Console.WriteLine("\n4. Updated style distribution:");
        var updatedStats = GetStyleDistribution(doc.Root);
        foreach (var (style, count) in updatedStats.OrderByDescending(x => x.Value))
        {
            Console.WriteLine($"   {style}: {count} paragraphs");
        }

        // 5. Save the modified document
        var outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_styles_modified.docx");

        Console.WriteLine($"\n5. Saving modified document to: {outputPath}");
        doc.SaveToFile(outputPath);

        // 6. Verify by re-parsing
        Console.WriteLine("\n6. Verifying saved document:");
        using var verifyParser = new WordDocumentTreeParser();
        var verifiedDoc = verifyParser.ParseFromFile(outputPath);

        var verifiedStats = GetStyleDistribution(verifiedDoc.Root);
        Console.WriteLine("   Style distribution in saved document:");
        foreach (var (style, count) in verifiedStats.OrderByDescending(x => x.Value))
        {
            Console.WriteLine($"   {style ?? "(no style)"}: {count} paragraphs");
        }

        Console.WriteLine("\n=== Demo Complete ===");
    }

    /// <summary>
    /// Finds all nodes with a specific paragraph style.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="styleId">The style ID to search for (e.g., "Heading1", "Normal", "Quote")</param>
    /// <returns>All nodes matching the specified style</returns>
    public static IEnumerable<DocumentNode> FindNodesByStyle(DocumentNode root, string styleId)
    {
        return GetAllNodes(root).Where(n =>
            n.ParagraphFormatting?.StyleId != null &&
            n.ParagraphFormatting.StyleId.Equals(styleId, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Changes the paragraph style of a node.
    /// Updates both the ParagraphFormatting.StyleId and the OriginalXml if present.
    /// </summary>
    /// <param name="node">The node to modify</param>
    /// <param name="newStyleId">The new style ID (e.g., "Heading2", "Quote", "NoSpacing")</param>
    public static void ChangeNodeStyle(DocumentNode node, string newStyleId)
    {
        // Ensure ParagraphFormatting exists
        node.ParagraphFormatting ??= new ParagraphFormatting();

        var oldStyleId = node.ParagraphFormatting.StyleId;

        // Update the StyleId
        node.ParagraphFormatting.StyleId = newStyleId;

        // If there's OriginalXml, update the style in the XML as well
        if (!string.IsNullOrEmpty(node.OriginalXml))
        {
            node.OriginalXml = UpdateStyleInXml(node.OriginalXml, oldStyleId, newStyleId);
        }

        // If this is a heading, update the HeadingLevel based on the new style
        if (newStyleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(newStyleId.Replace("Heading", ""), out var level))
        {
            node.HeadingLevel = level;
            node.Type = Core.ContentType.Heading;
        }
        else if (node.Type == Core.ContentType.Heading &&
                 !newStyleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
        {
            // Changing from a heading to a non-heading style
            node.HeadingLevel = 0;
            node.Type = Core.ContentType.Paragraph;
        }
    }

    /// <summary>
    /// Changes the style of multiple nodes matching a criteria.
    /// </summary>
    /// <param name="root">The root node to search from</param>
    /// <param name="fromStyleId">The style to search for</param>
    /// <param name="toStyleId">The style to change to</param>
    /// <returns>The number of nodes changed</returns>
    public static int ChangeStyleBulk(DocumentNode root, string fromStyleId, string toStyleId)
    {
        var nodes = FindNodesByStyle(root, fromStyleId).ToList();
        foreach (var node in nodes)
        {
            ChangeNodeStyle(node, toStyleId);
        }
        return nodes.Count;
    }

    /// <summary>
    /// Gets a dictionary of style IDs to their occurrence counts.
    /// </summary>
    public static Dictionary<string, int> GetStyleDistribution(DocumentNode root)
    {
        var distribution = new Dictionary<string, int>();

        foreach (var node in GetAllNodes(root))
        {
            if (node.Type is Core.ContentType.Paragraph or Core.ContentType.Heading or Core.ContentType.ListItem)
            {
                var styleId = node.ParagraphFormatting?.StyleId ?? "(no style)";
                distribution[styleId] = distribution.GetValueOrDefault(styleId, 0) + 1;
            }
        }

        return distribution;
    }

    /// <summary>
    /// Updates the style ID in an XML string.
    /// </summary>
    private static string UpdateStyleInXml(string xml, string? oldStyleId, string newStyleId)
    {
        // Pattern to match: <w:pStyle w:val="OldStyle"/>
        // This handles both self-closing and regular tags

        if (oldStyleId != null)
        {
            // Replace existing style
            var pattern = $@"<w:pStyle\s+w:val=""{Regex.Escape(oldStyleId)}""";
            var replacement = $@"<w:pStyle w:val=""{newStyleId}""";
            xml = Regex.Replace(xml, pattern, replacement, RegexOptions.IgnoreCase);
        }
        else
        {
            // No existing style - need to add one to the paragraph properties
            // Look for <w:pPr> and add the style inside it
            var pPrPattern = @"(<w:pPr[^>]*>)";
            var match = Regex.Match(xml, pPrPattern);
            if (match.Success)
            {
                var pPrTag = match.Groups[1].Value;
                xml = xml.Replace(pPrTag, $@"{pPrTag}<w:pStyle w:val=""{newStyleId}""/>");
            }
            else
            {
                // No pPr exists - need to add one after <w:p ...>
                var pTagPattern = @"(<w:p[^>]*>)";
                match = Regex.Match(xml, pTagPattern);
                if (match.Success)
                {
                    var pTag = match.Groups[1].Value;
                    xml = xml.Replace(pTag, $@"{pTag}<w:pPr><w:pStyle w:val=""{newStyleId}""/></w:pPr>");
                }
            }
        }

        return xml;
    }

    /// <summary>
    /// Gets all nodes in the tree (depth-first traversal).
    /// </summary>
    private static IEnumerable<DocumentNode> GetAllNodes(DocumentNode root)
    {
        yield return root;
        foreach (var child in root.Children)
        {
            foreach (var descendant in GetAllNodes(child))
            {
                yield return descendant;
            }
        }
    }

    /// <summary>
    /// Truncates a string to the specified length.
    /// </summary>
    private static string Truncate(string text, int maxLength)
    {
        text = text.Replace("\n", " ").Replace("\r", "").Trim();
        return text.Length <= maxLength ? text : text[..(maxLength - 3)] + "...";
    }
}
