using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;

namespace WordDocumentParser.Demo.Features.DocumentProperties;

/// <summary>
/// Demonstrates adding, changing, and removing a document property.
/// </summary>
public static class DocumentPropertyDemo
{
    private const string CustomPropertyFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    public static void Run(string inputPath)
    {
        var outputPathWithProperty = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_docprops_added.docx");
        var outputPathRemoved = Path.Combine(
            Path.GetDirectoryName(inputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_docprops_removed.docx");

        File.Copy(inputPath, outputPathWithProperty, true);

        const string propertyName = "DemoCustomProperty";

        Console.WriteLine($"Document 1 (with property): {outputPathWithProperty}");
        using (var doc = WordprocessingDocument.Open(outputPathWithProperty, true))
        {
            var customPart = doc.CustomFilePropertiesPart ?? doc.AddCustomFilePropertiesPart();
            customPart.Properties ??= new Properties();

            var props = customPart.Properties;

            Console.WriteLine("1. Adding custom property...");
            if (FindCustomProperty(props, propertyName) == null)
            {
                AddCustomProperty(props, propertyName, "Initial Value");
            }
            UpdateCustomProperty(props, propertyName, "Updated Value");
            props.Save();
            Console.WriteLine($"   Saved: {propertyName} = \"{GetCustomPropertyValue(props, propertyName) ?? "(missing)"}\"");
        }

        File.Copy(outputPathWithProperty, outputPathRemoved, true);
        Console.WriteLine($"Document 2 (property removed): {outputPathRemoved}");

        using (var doc = WordprocessingDocument.Open(outputPathRemoved, true))
        {
            var customPart = doc.CustomFilePropertiesPart ?? doc.AddCustomFilePropertiesPart();
            customPart.Properties ??= new Properties();
            var props = customPart.Properties;

            Console.WriteLine("2. Removing custom property...");
            if (RemoveCustomProperty(props, propertyName))
            {
                props.Save();
                Console.WriteLine($"   Removed: {propertyName}");
            }
            else
            {
                Console.WriteLine("   Property was not found to remove.");
            }
        }
    }

    private static CustomDocumentProperty? FindCustomProperty(Properties props, string name)
        => props.Elements<CustomDocumentProperty>()
            .FirstOrDefault(p => string.Equals(p.Name?.Value, name, StringComparison.OrdinalIgnoreCase));

    private static string? GetCustomPropertyValue(Properties props, string name)
    {
        var prop = FindCustomProperty(props, name);
        return prop?.InnerText;
    }

    private static void AddCustomProperty(Properties props, string name, string value)
    {
        var nextPid = props.Elements<CustomDocumentProperty>()
            .Select(p => (int?)p.PropertyId?.Value)
            .Max() ?? 1;

        var prop = new CustomDocumentProperty
        {
            Name = name,
            FormatId = CustomPropertyFormatId,
            PropertyId = nextPid + 1
        };

        prop.AppendChild(new VTLPWSTR(value));
        props.AppendChild(prop);
    }

    private static void UpdateCustomProperty(Properties props, string name, string value)
    {
        var prop = FindCustomProperty(props, name);
        if (prop == null)
        {
            AddCustomProperty(props, name, value);
            return;
        }

        prop.RemoveAllChildren();
        prop.AppendChild(new VTLPWSTR(value));
    }

    private static bool RemoveCustomProperty(Properties props, string name)
    {
        var prop = FindCustomProperty(props, name);
        if (prop == null)
        {
            return false;
        }

        prop.Remove();
        return true;
    }
}
