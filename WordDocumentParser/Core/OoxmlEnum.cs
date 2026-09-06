using DocumentFormat.OpenXml;

namespace WordDocumentParser.Core;

/// <summary>
/// Converts between the OpenXML SDK's typed enumeration attributes and the plain OOXML tokens
/// stored on this library's formatting models.
/// </summary>
/// <remarks>
/// <para>
/// In DocumentFormat.OpenXml 3.x the enumeration types are readonly structs, so
/// <c>JustificationValues.Center.ToString()</c> yields the record-struct rendering
/// <c>"JustificationValues { }"</c> rather than a usable name. Extracting formatting with
/// <c>Val?.Value.ToString()</c> therefore produced junk, and the writer's
/// <c>"Center" =&gt; JustificationValues.Center</c> switches never matched it — alignment,
/// break types, and every other enumerated property were silently dropped on the way out.
/// </para>
/// <para>
/// This helper reads and writes the wire token instead (<c>"center"</c>, <c>"page"</c>,
/// <c>"atLeast"</c>, …). Reader and writer stay symmetric by construction, and every value the
/// schema allows round-trips rather than only the handful a hand-written switch enumerated.
/// </para>
/// </remarks>
public static class OoxmlEnum
{
    /// <summary>
    /// Reads the OOXML token from a typed enumeration attribute, for example <c>"center"</c>.
    /// </summary>
    /// <typeparam name="T">The SDK enumeration struct.</typeparam>
    /// <param name="value">The attribute wrapper, which may be null or unset.</param>
    /// <returns>The wire token, or null when the attribute is absent.</returns>
    public static string? Token<T>(EnumValue<T>? value)
        where T : struct, IEnumValue, IEnumValueFactory<T>
        => value is null || !value.HasValue ? null : value.InnerText;

    /// <summary>
    /// Builds a typed enumeration attribute from an OOXML token.
    /// </summary>
    /// <remarks>
    /// Tokens are matched case-insensitively against the values the schema defines for
    /// <typeparamref name="T"/>, so both the wire form (<c>"atLeast"</c>) and the friendlier
    /// spellings callers tend to write by hand (<c>"AtLeast"</c>, <c>"atleast"</c>) are accepted.
    /// Unrecognised tokens return null rather than producing invalid XML.
    /// </remarks>
    /// <typeparam name="T">The SDK enumeration struct.</typeparam>
    /// <param name="token">The token to convert.</param>
    /// <returns>The typed attribute, or null when the token is empty or not valid for <typeparamref name="T"/>.</returns>
    public static EnumValue<T>? Parse<T>(string? token)
        where T : struct, IEnumValue, IEnumValueFactory<T>
    {
        if (string.IsNullOrWhiteSpace(token)) return null;

        var trimmed = token.Trim();
        return TryExact<T>(trimmed) ?? TryCaseInsensitive<T>(trimmed);
    }

    /// <summary>
    /// Converts an OOXML token to a typed enumeration value.
    /// </summary>
    /// <typeparam name="T">The SDK enumeration struct.</typeparam>
    /// <param name="token">The token to convert.</param>
    /// <param name="value">Receives the parsed value when the token is recognised.</param>
    /// <returns>True when the token maps to a value defined for <typeparamref name="T"/>.</returns>
    public static bool TryParseValue<T>(string? token, out T value)
        where T : struct, IEnumValue, IEnumValueFactory<T>
    {
        var parsed = Parse<T>(token);
        if (parsed is null)
        {
            value = default;
            return false;
        }

        value = parsed.Value;
        return true;
    }

    private static EnumValue<T>? TryExact<T>(string token)
        where T : struct, IEnumValue, IEnumValueFactory<T>
    {
        var candidate = new EnumValue<T> { InnerText = token };
        try
        {
            return ((IEnumValue)candidate.Value).IsValid ? candidate : null;
        }
        catch (FormatException)
        {
            return null;
        }
    }

    private static EnumValue<T>? TryCaseInsensitive<T>(string token)
        where T : struct, IEnumValue, IEnumValueFactory<T>
        => TokenMap<T>.ByToken.TryGetValue(token, out var value) ? new EnumValue<T>(value) : null;

    /// <summary>
    /// Caches the token-to-value map for one SDK enumeration. The SDK exposes its values as static
    /// properties rather than as an enumerable, so the map is built once per type by reflection.
    /// </summary>
    private static class TokenMap<T> where T : struct, IEnumValue, IEnumValueFactory<T>
    {
        public static readonly Dictionary<string, T> ByToken = Build();

        private static Dictionary<string, T> Build()
        {
            var map = new Dictionary<string, T>(StringComparer.OrdinalIgnoreCase);

            foreach (var property in typeof(T).GetProperties(
                         System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static))
            {
                if (property.PropertyType != typeof(T) || property.GetValue(null) is not T value) continue;

                var token = ((IEnumValue)value).Value;
                if (!string.IsNullOrEmpty(token))
                {
                    map[token] = value;
                }
            }

            return map;
        }
    }
}
