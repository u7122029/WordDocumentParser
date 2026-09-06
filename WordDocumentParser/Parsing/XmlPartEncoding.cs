using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using WordDocumentParser.Core;

namespace WordDocumentParser.Parsing;

/// <summary>Detects XML byte order and declared encoding before a character-bounded scan.</summary>
internal static class XmlPartEncoding
{
    /// <summary>
    /// Reads only the XML declaration. The full scan opens its own stream, so decoder read-ahead
    /// cannot lose bytes when the declaration selects a different encoding.
    /// </summary>
    public static Encoding Detect(OpenXmlPart part, long maxCharacters)
    {
        using var initial = part.GetStream(FileMode.Open, FileAccess.Read);
        Span<byte> prefix = stackalloc byte[4];
        var count = initial.ReadAtLeast(prefix, prefix.Length, throwOnEndOfStream: false);
        var bytes = prefix[..count];
        var encoding = DetectUnicode(bytes);
        var unicodeSignature = bytes.StartsWith(new byte[] { 0xEF, 0xBB, 0xBF }) ||
                               bytes.StartsWith(new byte[] { 0xFF, 0xFE }) ||
                               bytes.StartsWith(new byte[] { 0xFE, 0xFF }) ||
                               bytes.Contains((byte)0);

        using var probe = initial.CanSeek ? initial : part.GetStream(FileMode.Open, FileAccess.Read);
        if (probe.CanSeek) probe.Position = 0;
        // Replacement fallback is intentional only for the probe: its buffer may extend past an
        // ASCII declaration into text whose encoding has not been selected yet.
        using var reader = new StreamReader(probe, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 128);
        Span<char> start = stackalloc char[6];
        var read = reader.ReadBlock(start);
        encoding = reader.CurrentEncoding;
        if (read != start.Length || !start[..5].SequenceEqual("<?xml") || !char.IsWhiteSpace(start[5]))
            return Strict(encoding);

        var declaration = new StringBuilder();
        declaration.Append(start);
        while (true)
        {
            if (maxCharacters > 0 && declaration.Length > maxCharacters)
                throw new DocumentLimitExceededException(
                    $"Document exceeds the limit of {maxCharacters} characters in a part (reading {part.Uri}).");
            var value = reader.Read();
            if (value < 0) throw new XmlException("Unterminated XML declaration.");
            declaration.Append((char)value);
            if (value == '>' && declaration[^2] == '?') break;
        }

        using var xml = XmlReader.Create(new StringReader(declaration.ToString()), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null
        });
        xml.Read();
        var name = xml.GetAttribute("encoding");
        if (name is null) return Strict(encoding);

        Encoding declared;
        try { declared = Encoding.GetEncoding(name); }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
        {
            throw new XmlException($"Unsupported XML encoding '{name}'.", ex);
        }

        if (!unicodeSignature) return Strict(declared);
        // An unqualified UTF-16/UTF-32 declaration lets the signature choose byte order.
        if (declared.CodePage == encoding.CodePage ||
            (name.Equals("utf-16", StringComparison.OrdinalIgnoreCase) && encoding.CodePage is 1200 or 1201) ||
            (name.Equals("utf-32", StringComparison.OrdinalIgnoreCase) && encoding.CodePage is 12000 or 12001))
            return Strict(encoding);
        throw new XmlException("The XML encoding declaration conflicts with its byte-order signature.");
    }

    private static Encoding Strict(Encoding encoding) => Encoding.GetEncoding(
        encoding.CodePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

    private static Encoding DetectUnicode(ReadOnlySpan<byte> bytes)
    {
        if (bytes.StartsWith(new byte[] { 0x00, 0x00, 0xFE, 0xFF }) ||
            bytes.StartsWith(new byte[] { 0x00, 0x00, 0x00, 0x3C })) return new UTF32Encoding(true, false);
        if (bytes.StartsWith(new byte[] { 0xFF, 0xFE, 0x00, 0x00 }) ||
            bytes.StartsWith(new byte[] { 0x3C, 0x00, 0x00, 0x00 })) return Encoding.UTF32;
        if (bytes.StartsWith(new byte[] { 0xFE, 0xFF }) ||
            bytes.StartsWith(new byte[] { 0x00, 0x3C })) return Encoding.BigEndianUnicode;
        if (bytes.StartsWith(new byte[] { 0xFF, 0xFE }) ||
            bytes.StartsWith(new byte[] { 0x3C, 0x00 })) return Encoding.Unicode;
        return Encoding.UTF8;
    }
}
