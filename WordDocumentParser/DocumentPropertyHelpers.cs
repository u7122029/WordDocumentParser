using WordDocumentParser.Models.ContentControls;

namespace WordDocumentParser;

internal static class DocumentPropertyHelpers
{
    private static readonly HashSet<string> CorePropertyNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "title", "subject", "creator", "author", "keywords", "description", "comments",
        "lastmodifiedby", "revision", "created", "modified", "category", "contentstatus", "status"
    };

    private static readonly HashSet<string> ExtendedPropertyNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "template", "application", "appversion", "company", "manager", "pages", "words",
        "characters", "characterswithspaces", "lines", "paragraphs", "totaltime"
    };

    internal static bool IsCoreProperty(string propertyName) => CorePropertyNames.Contains(propertyName);

    internal static bool IsExtendedProperty(string propertyName) => ExtendedPropertyNames.Contains(propertyName);

    internal static DocumentPropertyType DeterminePropertyType(string propertyName)
    {
        if (IsCoreProperty(propertyName))
            return DocumentPropertyType.Core;
        if (IsExtendedProperty(propertyName))
            return DocumentPropertyType.Extended;
        return DocumentPropertyType.Custom;
    }

    internal static string ExtractPropertyNameFromXPath(string xpath)
    {
        var parts = xpath.Split('/');
        if (parts.Length == 0) return xpath;

        var lastPart = parts[^1];
        var colonIndex = lastPart.IndexOf(':');
        if (colonIndex >= 0)
            lastPart = lastPart[(colonIndex + 1)..];
        var bracketIndex = lastPart.IndexOf('[');
        if (bracketIndex >= 0)
            lastPart = lastPart[..bracketIndex];
        return lastPart;
    }
}
