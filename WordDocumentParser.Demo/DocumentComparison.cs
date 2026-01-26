using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.CustomProperties;

namespace WordDocumentParser.Demo;

/// <summary>
/// Provides document comparison functionality for demos.
/// </summary>
public static class DocumentComparison
{
    /// <summary>
    /// Compares two Word documents and reports differences to the console.
    /// </summary>
    public static void CompareDocuments(string originalPath, string copyPath)
    {
        using var origDoc = WordprocessingDocument.Open(originalPath, false);
        using var copyDoc = WordprocessingDocument.Open(copyPath, false);

        CompareDocumentProperties(origDoc, copyDoc);
        CompareDynamicReferences(origDoc, copyDoc);
        CompareGlossaryDocument(origDoc, copyDoc);
        CompareStructure(origDoc, copyDoc);
    }

    private static void CompareDocumentProperties(
        WordprocessingDocument origDoc,
        WordprocessingDocument copyDoc)
    {
        Console.WriteLine("\n--- Document Properties ---");

        // Core properties
        var origProps = origDoc.PackageProperties;
        var copyProps = copyDoc.PackageProperties;

        Console.WriteLine("Core Properties:");
        CompareProperty("Title", origProps.Title, copyProps.Title);
        CompareProperty("Subject", origProps.Subject, copyProps.Subject);
        CompareProperty("Creator", origProps.Creator, copyProps.Creator);
        CompareProperty("Keywords", origProps.Keywords, copyProps.Keywords);
        CompareProperty("Category", origProps.Category, copyProps.Category);

        // Extended properties
        Console.WriteLine("\nExtended Properties:");
        var origExtProps = origDoc.ExtendedFilePropertiesPart?.Properties;
        var copyExtProps = copyDoc.ExtendedFilePropertiesPart?.Properties;

        if (origExtProps != null || copyExtProps != null)
        {
            CompareProperty("Company", origExtProps?.Company?.Text, copyExtProps?.Company?.Text);
            CompareProperty("Template", origExtProps?.Template?.Text, copyExtProps?.Template?.Text);
            CompareProperty("Manager", origExtProps?.Manager?.Text, copyExtProps?.Manager?.Text);
        }
        else
        {
            Console.WriteLine("  (No extended properties found)");
        }

        // Custom properties
        Console.WriteLine("\nCustom Properties:");
        var origCustomProps = origDoc.CustomFilePropertiesPart?.Properties;
        var copyCustomProps = copyDoc.CustomFilePropertiesPart?.Properties;

        if (origCustomProps != null)
        {
            var origCount = origCustomProps.Elements<CustomDocumentProperty>().Count();
            var copyCount = copyCustomProps?.Elements<CustomDocumentProperty>().Count() ?? 0;
            Console.WriteLine($"  Count: Original={origCount}, Copy={copyCount} {StatusText(origCount == copyCount)}");

            foreach (var prop in origCustomProps.Elements<CustomDocumentProperty>())
            {
                Console.WriteLine($"    {prop.Name?.Value ?? "(unnamed)"}: \"{prop.InnerText ?? "(null)"}\"");
            }
        }
        else
        {
            Console.WriteLine("  (No custom properties in original)");
        }
    }

    private static void CompareDynamicReferences(
        WordprocessingDocument origDoc,
        WordprocessingDocument copyDoc)
    {
        Console.WriteLine("\n--- Dynamic Reference Fields ---");

        var origBody = origDoc.MainDocumentPart?.Document?.Body;
        var copyBody = copyDoc.MainDocumentPart?.Document?.Body;

        if (origBody == null || copyBody == null)
        {
            Console.WriteLine("Could not access document bodies");
            return;
        }

        // Find DOCPROPERTY field codes
        var origDocPropFields = origBody.Descendants<FieldCode>()
            .Where(f => f.Text?.Contains("DOCPROPERTY") == true)
            .Select(f => f.Text?.Trim())
            .ToList();

        var copyDocPropFields = copyBody.Descendants<FieldCode>()
            .Where(f => f.Text?.Contains("DOCPROPERTY") == true)
            .Select(f => f.Text?.Trim())
            .ToList();

        Console.WriteLine($"DOCPROPERTY fields: Original={origDocPropFields.Count}, Copy={copyDocPropFields.Count}");
        foreach (var field in origDocPropFields)
        {
            var existsInCopy = copyDocPropFields.Contains(field);
            Console.WriteLine($"  {field} {StatusText(existsInCopy)}");
        }

        // All field codes summary
        var allOrigFields = origBody.Descendants<FieldCode>()
            .Select(f => f.Text?.Trim())
            .Where(t => !string.IsNullOrEmpty(t))
            .ToList();

        var allCopyFields = copyBody.Descendants<FieldCode>()
            .Select(f => f.Text?.Trim())
            .Where(t => !string.IsNullOrEmpty(t))
            .ToList();

        Console.WriteLine($"\nAll field codes: Original={allOrigFields.Count}, Copy={allCopyFields.Count}");
    }

    private static void CompareGlossaryDocument(
        WordprocessingDocument origDoc,
        WordprocessingDocument copyDoc)
    {
        Console.WriteLine("\n--- Glossary Document ---");

        var origGlossary = origDoc.MainDocumentPart?.GlossaryDocumentPart;
        var copyGlossary = copyDoc.MainDocumentPart?.GlossaryDocumentPart;

        var origHasGlossary = origGlossary?.GlossaryDocument != null;
        var copyHasGlossary = copyGlossary?.GlossaryDocument != null;

        Console.WriteLine($"Has Glossary: Original={origHasGlossary}, Copy={copyHasGlossary} {StatusText(origHasGlossary == copyHasGlossary)}");

        if (origHasGlossary)
        {
            var origDocPartCount = origGlossary!.GlossaryDocument!.Descendants<DocPart>().Count();
            var copyDocPartCount = copyGlossary?.GlossaryDocument?.Descendants<DocPart>().Count() ?? 0;
            Console.WriteLine($"Building blocks: Original={origDocPartCount}, Copy={copyDocPartCount} {StatusText(origDocPartCount == copyDocPartCount)}");
        }
    }

    private static void CompareStructure(
        WordprocessingDocument origDoc,
        WordprocessingDocument copyDoc)
    {
        Console.WriteLine("\n--- Document Structure ---");

        var origBody = origDoc.MainDocumentPart?.Document?.Body;
        var copyBody = copyDoc.MainDocumentPart?.Document?.Body;

        if (origBody == null || copyBody == null) return;

        var origParas = origBody.Descendants<Paragraph>().Count();
        var copyParas = copyBody.Descendants<Paragraph>().Count();
        Console.WriteLine($"Paragraphs: Original={origParas}, Copy={copyParas} {StatusText(origParas == copyParas)}");

        var origTables = origBody.Descendants<Table>().Count();
        var copyTables = copyBody.Descendants<Table>().Count();
        Console.WriteLine($"Tables: Original={origTables}, Copy={copyTables} {StatusText(origTables == copyTables)}");

        var origFldChar = origBody.Descendants<FieldChar>().Count();
        var copyFldChar = copyBody.Descendants<FieldChar>().Count();
        Console.WriteLine($"Field characters: Original={origFldChar}, Copy={copyFldChar} {StatusText(origFldChar == copyFldChar)}");

        var origSectPr = origBody.Descendants<SectionProperties>().Count();
        var copySectPr = copyBody.Descendants<SectionProperties>().Count();
        Console.WriteLine($"Section properties: Original={origSectPr}, Copy={copySectPr} {StatusText(origSectPr == copySectPr)}");
    }

    private static void CompareProperty(string name, string? original, string? copy)
    {
        var origDisplay = original ?? "(null)";
        var copyDisplay = copy ?? "(null)";
        Console.WriteLine($"  {name}: Original=\"{origDisplay}\" | Copy=\"{copyDisplay}\" {StatusText(original == copy)}");
    }

    private static string StatusText(bool match) => match ? "[OK]" : "[DIFF]";
}