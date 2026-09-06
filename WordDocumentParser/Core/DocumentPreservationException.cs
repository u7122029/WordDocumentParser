namespace WordDocumentParser.Core;

/// <summary>
/// Thrown when a part, relationship, or property of the source document could not be preserved.
/// </summary>
/// <remarks>
/// Preservation failures surface by default, so a lossy result is never mistaken for a faithful
/// one. Callers who would rather salvage what they can set
/// <see cref="RecoveryOptions.ContinueOnPreservationFailure"/> and read
/// <see cref="RecoveryOptions.Diagnostics"/> afterwards.
/// </remarks>
public class DocumentPreservationException : InvalidOperationException
{
    /// <summary>Creates the exception for a named part or property.</summary>
    /// <param name="target">The part URI, relationship ID, or property name that failed.</param>
    /// <param name="message">What went wrong.</param>
    /// <param name="innerException">The underlying failure, when there was one.</param>
    public DocumentPreservationException(string target, string message, Exception? innerException = null)
        : base($"{message} (target: {target})", innerException)
        => Target = target;

    /// <summary>The part URI, relationship ID, or property name that failed.</summary>
    public string Target { get; }
}

/// <summary>
/// One recorded preservation failure.
/// </summary>
/// <param name="Target">The part URI, relationship ID, or property name that failed.</param>
/// <param name="Message">What went wrong.</param>
/// <param name="Exception">The underlying failure, when there was one.</param>
public sealed record PreservationDiagnostic(string Target, string Message, Exception? Exception = null)
{
    /// <summary>Returns a single-line description of the failure.</summary>
    public override string ToString() =>
        Exception is null ? $"{Target}: {Message}" : $"{Target}: {Message} ({Exception.GetType().Name})";
}

/// <summary>
/// Controls what happens when part of a document cannot be preserved.
/// </summary>
public sealed class RecoveryOptions
{
    /// <summary>
    /// When true, preservation failures are recorded in <see cref="Diagnostics"/> and processing
    /// continues with that piece of content missing. When false (the default) the failure is thrown
    /// as a <see cref="DocumentPreservationException"/>, so a lossy result is never mistaken for a
    /// faithful one.
    /// </summary>
    public bool ContinueOnPreservationFailure { get; init; }

    /// <summary>
    /// Failures recorded during the operation. Only populated when
    /// <see cref="ContinueOnPreservationFailure"/> is true.
    /// </summary>
    public List<PreservationDiagnostic> Diagnostics { get; } = [];

    /// <summary>True when at least one failure was recorded.</summary>
    public bool HasDiagnostics => Diagnostics.Count > 0;

    /// <summary>
    /// Records a failure, or rethrows it as a <see cref="DocumentPreservationException"/> when
    /// recovery is not enabled.
    /// </summary>
    /// <param name="target">The part URI, relationship ID, or property name that failed.</param>
    /// <param name="message">What went wrong.</param>
    /// <param name="exception">The underlying failure, when there was one.</param>
    internal void Report(string target, string message, Exception? exception = null)
    {
        if (!ContinueOnPreservationFailure)
        {
            throw new DocumentPreservationException(target, message, exception);
        }

        Diagnostics.Add(new PreservationDiagnostic(target, message, exception));
    }
}
