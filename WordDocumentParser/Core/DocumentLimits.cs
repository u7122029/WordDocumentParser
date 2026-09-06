namespace WordDocumentParser.Core;

/// <summary>
/// Bounds applied while parsing a document, so that a hostile or malformed <c>.docx</c> cannot
/// exhaust memory on the parsing host.
/// </summary>
/// <remarks>
/// <para>
/// A <c>.docx</c> is a zip archive, and both its XML parts and its media parts decompress to an
/// attacker-chosen size. The OpenXML SDK ships with <c>MaxCharactersInPart</c> defaulted to zero —
/// that is, unlimited — and Microsoft documents raising it as the denial-of-service mitigation for
/// untrusted input. This library additionally reads every media part into memory, so an aggregate
/// budget is needed on top of the SDK's per-part character cap.
/// </para>
/// <para>
/// <see cref="Trusted"/> keeps the historical unbounded behaviour and is the default, so existing
/// callers who parse their own files are unaffected. Anything parsing uploads should pass
/// <see cref="Untrusted"/> or a tuned instance.
/// </para>
/// </remarks>
public sealed class DocumentLimits
{
    /// <summary>
    /// No limits. The default, appropriate when the caller controls the input files.
    /// </summary>
    public static DocumentLimits Trusted { get; } = new();

    /// <summary>
    /// Conservative limits for documents arriving from outside the trust boundary:
    /// 32M characters per XML part, 256 MiB of decompressed binary parts, 4096 parts,
    /// and 100 levels of XML element nesting.
    /// </summary>
    public static DocumentLimits Untrusted { get; } = new()
    {
        MaxCharactersInPart = 32L * 1024 * 1024,
        MaxTotalBinaryBytes = 256L * 1024 * 1024,
        MaxPartCount = 4096,
        MaxElementDepth = 100
    };

    /// <summary>
    /// Maximum characters the SDK will read from any single XML part, or 0 for unlimited.
    /// Passed through to <c>OpenSettings.MaxCharactersInPart</c>.
    /// </summary>
    public long MaxCharactersInPart { get; init; }

    /// <summary>
    /// Maximum total decompressed size of the binary parts (images and other media) the parser will
    /// hold in memory, or 0 for unlimited.
    /// </summary>
    public long MaxTotalBinaryBytes { get; init; }

    /// <summary>
    /// Maximum number of package parts the parser will read, or 0 for unlimited.
    /// </summary>
    public int MaxPartCount { get; init; }

    /// <summary>
    /// Maximum XML element nesting depth accepted in any part, or 0 for unlimited.
    /// </summary>
    /// <remarks>
    /// This counts raw XML elements, not document structures: an ordinary paragraph already sits
    /// four levels deep, and each nested table adds roughly six. The OpenXML SDK builds its object
    /// model by recursive descent, so this is what stops a deeply nested part from overflowing the
    /// stack — a failure no caller can catch. 100 leaves ample room for real documents.
    /// </remarks>
    public int MaxElementDepth { get; init; }

    /// <summary>
    /// Throws when a running total of decompressed binary bytes exceeds <see cref="MaxTotalBinaryBytes"/>.
    /// </summary>
    /// <param name="totalBytes">Bytes read so far, including the part being added.</param>
    /// <exception cref="DocumentLimitExceededException">The budget is exhausted.</exception>
    public void EnforceBinaryBudget(long totalBytes)
    {
        if (MaxTotalBinaryBytes > 0 && totalBytes > MaxTotalBinaryBytes)
        {
            throw new DocumentLimitExceededException(
                $"Document exceeds the binary part budget of {MaxTotalBinaryBytes} bytes.");
        }
    }

    /// <summary>
    /// Throws when the number of parts read exceeds <see cref="MaxPartCount"/>.
    /// </summary>
    /// <param name="partCount">Parts read so far, including the part being added.</param>
    /// <exception cref="DocumentLimitExceededException">The budget is exhausted.</exception>
    public void EnforcePartCount(int partCount)
    {
        if (MaxPartCount > 0 && partCount > MaxPartCount)
        {
            throw new DocumentLimitExceededException(
                $"Document exceeds the limit of {MaxPartCount} package parts.");
        }
    }

    /// <summary>
    /// Throws when element nesting exceeds <see cref="MaxElementDepth"/>.
    /// </summary>
    /// <param name="depth">The depth about to be entered.</param>
    /// <exception cref="DocumentLimitExceededException">The budget is exhausted.</exception>
    public void EnforceDepth(int depth)
    {
        if (MaxElementDepth > 0 && depth > MaxElementDepth)
        {
            throw new DocumentLimitExceededException(
                $"Document exceeds the maximum element nesting depth of {MaxElementDepth}.");
        }
    }
}

/// <summary>
/// Thrown when a document exceeds one of the bounds in <see cref="DocumentLimits"/>.
/// </summary>
public class DocumentLimitExceededException : InvalidOperationException
{
    /// <summary>Creates the exception with a message describing the exceeded bound.</summary>
    /// <param name="message">The message.</param>
    public DocumentLimitExceededException(string message) : base(message) { }

    /// <summary>Creates the exception with a message and an inner cause.</summary>
    /// <param name="message">The message.</param>
    /// <param name="innerException">The underlying failure.</param>
    public DocumentLimitExceededException(string message, Exception innerException)
        : base(message, innerException) { }
}
