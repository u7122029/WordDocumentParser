using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordDocumentParser.Extensions;

/// <summary>
/// Merges one document's numbering definitions into another's, renumbering to avoid collisions.
/// </summary>
/// <remarks>
/// A list paragraph carries only a <c>w:numId</c>; the definition that gives it its bullet or number
/// sequence lives in <c>numbering.xml</c>. Appending a document without bringing its numbering along
/// left those references dangling — the list either lost its formatting or silently adopted whatever
/// definition the target happened to have under the same ID.
/// </remarks>
internal static class NumberingMerge
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// Copies the source document's numbering definitions into the target.
    /// </summary>
    /// <param name="target">The document receiving the definitions.</param>
    /// <param name="source">The document supplying them.</param>
    /// <returns>
    /// A map from the source's numbering IDs to the IDs they were given in the target, empty when
    /// there was nothing to merge.
    /// </returns>
    public static Dictionary<int, int> Merge(WordDocument target, WordDocument source)
    {
        var mapping = new Dictionary<int, int>();

        if (string.IsNullOrEmpty(source.PackageData.NumberingXml))
        {
            return mapping;
        }

        XElement sourceRoot;
        try
        {
            sourceRoot = XElement.Parse(source.PackageData.NumberingXml);
        }
        catch (System.Xml.XmlException)
        {
            return mapping;
        }

        var targetRoot = ParseOrCreateTargetNumbering(target);

        // IDs are kept where the target is not already using them, so merging into a document with
        // no lists of its own leaves the source's numbering exactly as it was.
        var usedAbstractIds = CollectIds(targetRoot, "abstractNum", "abstractNumId");
        var usedNumIds = CollectIds(targetRoot, "num", "numId");

        // Remap the abstract definitions first: concrete instances point at them by ID.
        var abstractMapping = new Dictionary<int, int>();
        foreach (var abstractNum in sourceRoot.Elements(W + "abstractNum"))
        {
            if (!TryGetAttributeValue(abstractNum, "abstractNumId", out var originalId)) continue;

            var copy = new XElement(abstractNum);
            var newId = AllocateId(usedAbstractIds, originalId);
            copy.SetAttributeValue(W + "abstractNumId", newId);

            // nsid and tmpl identify a definition across documents; dropping them lets Word treat the
            // copy as its own definition rather than reconciling it with the target's.
            copy.Elements(W + "nsid").Remove();
            copy.Elements(W + "tmpl").Remove();

            abstractMapping[originalId] = newId;
            targetRoot.Add(copy);
        }

        foreach (var num in sourceRoot.Elements(W + "num"))
        {
            if (!TryGetAttributeValue(num, "numId", out var originalId)) continue;

            var copy = new XElement(num);
            var newId = AllocateId(usedNumIds, originalId);
            copy.SetAttributeValue(W + "numId", newId);

            var abstractRef = copy.Element(W + "abstractNumId");
            if (abstractRef is not null &&
                int.TryParse(abstractRef.Attribute(W + "val")?.Value, out var abstractId) &&
                abstractMapping.TryGetValue(abstractId, out var newAbstractId))
            {
                abstractRef.SetAttributeValue(W + "val", newAbstractId);
            }

            if (newId != originalId)
            {
                mapping[originalId] = newId;
            }

            targetRoot.Add(copy);
        }

        if (abstractMapping.Count > 0 || sourceRoot.Elements(W + "num").Any())
        {
            SortNumberingChildren(targetRoot);
            target.PackageData.NumberingXml = targetRoot.ToString();
        }

        return mapping;
    }

    /// <summary>
    /// Keeps the preferred ID when it is free, otherwise takes the next unused one.
    /// </summary>
    private static int AllocateId(HashSet<int> used, int preferred)
    {
        if (used.Add(preferred))
        {
            return preferred;
        }

        var candidate = used.Count == 0 ? 1 : used.Max() + 1;
        while (!used.Add(candidate))
        {
            candidate++;
        }

        return candidate;
    }

    private static HashSet<int> CollectIds(XElement root, string elementName, string attributeName) =>
        [.. root.Elements(W + elementName)
            .Select(e => int.TryParse(e.Attribute(W + attributeName)?.Value, out var value) ? value : -1)
            .Where(value => value >= 0)];

    /// <summary>
    /// Matches a <c>numId</c> element's value attribute, whatever prefix or spacing it uses.
    /// The <c>\b</c> after the element name keeps longer names such as <c>numIdMacAtCleanup</c> out.
    /// </summary>
    private static readonly Regex NumberingReference = new(
        @"<(?:[\w.\-]+:)?numId\b[^>]*?\b(?:[\w.\-]+:)?val=""(?<value>-?\d+)""",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// Rewrites <c>w:numId</c> references in a node's XML to the IDs the merge assigned.
    /// </summary>
    /// <param name="xml">The node XML, which may be null.</param>
    /// <param name="mapping">The numbering ID map.</param>
    /// <returns>The rewritten XML.</returns>
    /// <remarks>
    /// Each reference is matched once and its original value looked up once. Replacing the ids one
    /// after another instead let a later substitution rewrite what an earlier one had just written —
    /// mapping 1 to 2 and then 2 to 3 collapsed two independent lists onto the same definition.
    /// </remarks>
    public static string? ApplyMapping(string? xml, Dictionary<int, int> mapping)
    {
        if (string.IsNullOrEmpty(xml) || mapping.Count == 0) return xml;

        return NumberingReference.Replace(xml, match =>
        {
            var value = match.Groups["value"];

            if (!int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ||
                !mapping.TryGetValue(id, out var newId))
            {
                return match.Value;
            }

            var start = value.Index - match.Index;
            return string.Concat(
                match.Value.AsSpan(0, start),
                newId.ToString(CultureInfo.InvariantCulture),
                match.Value.AsSpan(start + value.Length));
        });
    }

    private static XElement ParseOrCreateTargetNumbering(WordDocument target)
    {
        if (!string.IsNullOrEmpty(target.PackageData.NumberingXml))
        {
            try
            {
                return XElement.Parse(target.PackageData.NumberingXml);
            }
            catch (System.Xml.XmlException)
            {
                // Fall through and start from an empty definition rather than corrupting the merge.
            }
        }

        return new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W));
    }

    private static bool TryGetAttributeValue(XElement element, string attributeName, out int value) =>
        int.TryParse(element.Attribute(W + attributeName)?.Value, out value);

    /// <summary>
    /// Restores the schema's required ordering: every <c>abstractNum</c> precedes every <c>num</c>.
    /// </summary>
    private static void SortNumberingChildren(XElement root)
    {
        var ordered = root.Elements()
            .OrderBy(e => e.Name.LocalName switch
            {
                "numPicBullet" => 0,
                "abstractNum" => 1,
                "num" => 2,
                _ => 3
            })
            .ToList();

        root.ReplaceNodes(ordered);
    }
}
