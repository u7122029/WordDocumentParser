using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;

namespace WordDocumentParser.Parsing;

/// <summary>
/// Holds shared state during document parsing.
/// Passed to extractors to avoid tight coupling.
/// </summary>
internal sealed class ParsingContext
{
    public required WordprocessingDocument Document { get; init; }
    public required MainDocumentPart MainPart { get; init; }
    public required DocumentLimits Limits { get; init; }
    public required RecoveryOptions Recovery { get; init; }

    public Dictionary<string, Style> StyleCache { get; } = [];
    public Dictionary<string, string> HyperlinkUrls { get; } = [];

    /// <summary>
    /// Decompressed media parts keyed by part URI, so a part referenced many times is read and held
    /// once rather than once per reference.
    /// </summary>
    private readonly Dictionary<string, byte[]> _mediaCache = [];

    /// <summary>Running total of decompressed media bytes held, for the binary budget.</summary>
    private long _mediaBytes;

    /// <summary>
    /// Reads a binary part, returning the shared buffer if this part has already been read.
    /// </summary>
    /// <param name="part">The part to read.</param>
    /// <returns>The part's bytes. Callers must treat the array as immutable.</returns>
    /// <exception cref="DocumentLimitExceededException">The binary budget is exhausted.</exception>
    /// <remarks>
    /// The budget is checked as the part is read, not after. A part is decompressed by the act of
    /// reading it, so copying it whole and only then comparing against the limit lets a hostile
    /// document exhaust memory before it can be rejected.
    /// </remarks>
    public byte[] ReadBinaryPart(OpenXmlPart part)
    {
        var key = part.Uri.ToString();
        if (_mediaCache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        using var stream = part.GetStream();
        var bytes = ReadBounded(stream, Limits, _mediaBytes, part.Uri.ToString());

        _mediaBytes += bytes.Length;
        _mediaCache[key] = bytes;
        return bytes;
    }

    /// <summary>
    /// Copies a stream into memory, stopping as soon as the aggregate binary budget is exceeded.
    /// </summary>
    /// <param name="stream">The stream to read.</param>
    /// <param name="limits">The limits to enforce.</param>
    /// <param name="alreadyRead">Bytes already counted against the budget.</param>
    /// <param name="target">What is being read, for the error message.</param>
    /// <returns>The bytes read.</returns>
    /// <exception cref="DocumentLimitExceededException">The budget is exhausted.</exception>
    internal static byte[] ReadBounded(Stream stream, DocumentLimits limits, long alreadyRead, string target)
    {
        var budget = limits.MaxTotalBinaryBytes;
        using var buffer = new MemoryStream();
        var chunk = new byte[81920];

        while (true)
        {
            var read = stream.Read(chunk, 0, chunk.Length);
            if (read == 0) break;

            buffer.Write(chunk, 0, read);

            if (budget > 0 && alreadyRead + buffer.Length > budget)
            {
                throw new DocumentLimitExceededException(
                    $"Document exceeds the binary part budget of {budget} bytes (reading {target}).");
            }
        }

        return buffer.ToArray();
    }

    /// <summary>
    /// Reads an XML part as text, stopping as soon as the per-part character limit is exceeded.
    /// </summary>
    /// <param name="part">The part to read.</param>
    /// <returns>The part's text.</returns>
    /// <exception cref="DocumentLimitExceededException">The part is longer than the limit allows.</exception>
    /// <remarks>
    /// The SDK's own <c>MaxCharactersInPart</c> only covers the parts it parses. Parts this library
    /// reads as raw text — custom XML, the property parts — bypassed it entirely, so the advertised
    /// limit did not bound them at all.
    /// </remarks>
    public string ReadTextPart(OpenXmlPart part)
    {
        using var stream = part.GetStream();
        return ReadBoundedText(stream, Limits, part.Uri.ToString());
    }

    /// <summary>
    /// Reads a stream as text, stopping as soon as the per-part character limit is exceeded.
    /// </summary>
    /// <param name="stream">The stream to read.</param>
    /// <param name="limits">The limits to enforce.</param>
    /// <param name="target">What is being read, for the error message.</param>
    /// <returns>The text read.</returns>
    /// <exception cref="DocumentLimitExceededException">The limit is exceeded.</exception>
    internal static string ReadBoundedText(Stream stream, DocumentLimits limits, string target)
    {
        var limit = limits.MaxCharactersInPart;
        using var reader = new StreamReader(stream);

        if (limit <= 0)
        {
            return reader.ReadToEnd();
        }

        var builder = new System.Text.StringBuilder();
        var chunk = new char[8192];

        while (true)
        {
            var read = reader.Read(chunk, 0, chunk.Length);
            if (read == 0) break;

            builder.Append(chunk, 0, read);

            if (builder.Length > limit)
            {
                throw new DocumentLimitExceededException(
                    $"Document exceeds the limit of {limit} characters in a part (reading {target}).");
            }
        }

        return builder.ToString();
    }


    /// <summary>
    /// Caches styles from the document for quick lookup.
    /// </summary>
    public void CacheStyles()
    {
        var stylesPart = MainPart.StyleDefinitionsPart;
        if (stylesPart?.Styles is null) return;

        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            if (style.StyleId?.Value is not null)
            {
                StyleCache[style.StyleId.Value] = style;
            }
        }
    }

    /// <summary>
    /// Caches hyperlink relationships for URL resolution.
    /// </summary>
    public void CacheHyperlinkRelationships()
    {
        foreach (var rel in MainPart.HyperlinkRelationships)
        {
            HyperlinkUrls[rel.Id] = rel.Uri.ToString();
        }
    }

    /// <summary>
    /// Gets the URL for a hyperlink relationship ID.
    /// </summary>
    public string? GetHyperlinkUrl(string? relationshipId)
        => relationshipId is not null && HyperlinkUrls.TryGetValue(relationshipId, out var url) ? url : null;

    /// <summary>
    /// Gets a style by its ID.
    /// </summary>
    public Style? GetStyle(string? styleId)
        => styleId is not null && StyleCache.TryGetValue(styleId, out var style) ? style : null;
}
