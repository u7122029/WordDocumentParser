using WordDocumentParser.Core;
using WordDocumentParser.Writing;

namespace WordDocumentParser;

/// <summary>
/// Writes a document tree structure to a Word document (.docx file).
/// </summary>
/// <remarks>
/// <para>
/// When the document was parsed with the source package retained, the writer edits a copy of that
/// package: it replaces the body with the current tree and rewrites only the parts a caller changed.
/// Parts this library has no model for pass through untouched, and an unedited document round-trips
/// with its content intact.
/// </para>
/// <para>
/// A writer instance holds no state between calls, so it can be reused sequentially. It is not safe
/// to share one instance across threads; use one per thread.
/// </para>
/// </remarks>
public class WordDocumentTreeWriter : IDocumentWriter
{
    /// <summary>
    /// What to do when part of the document cannot be preserved. By default such failures throw, so
    /// a lossy save is never reported as a successful one.
    /// </summary>
    public RecoveryOptions Recovery { get; init; } = new();

    /// <summary>
    /// Writes a document to a file, replacing any existing file only once the new one is complete.
    /// </summary>
    /// <param name="document">The document to write.</param>
    /// <param name="filePath">The destination path.</param>
    /// <remarks>
    /// The package is built in memory and staged in a sibling temporary file, which then replaces
    /// the destination. Creating the destination first and building into it meant a serialization
    /// failure destroyed whatever was already there.
    /// </remarks>
    public void WriteToFile(WordDocument document, string filePath)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentException.ThrowIfNullOrEmpty(filePath);

        var bytes = BuildPackage(document);

        var directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var stagingPath = $"{filePath}.{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllBytes(stagingPath, bytes);
            File.Move(stagingPath, filePath, overwrite: true);
        }
        finally
        {
            if (File.Exists(stagingPath))
            {
                File.Delete(stagingPath);
            }
        }
    }

    /// <summary>
    /// Writes a document to a stream.
    /// </summary>
    /// <param name="document">The document to write.</param>
    /// <param name="stream">The destination stream, which is written to but not disposed.</param>
    /// <remarks>
    /// The package is assembled fully before the stream is touched, so a failure leaves the stream
    /// unwritten rather than holding a partial package.
    /// </remarks>
    public void WriteToStream(WordDocument document, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(stream);

        var bytes = BuildPackage(document);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>
    /// Builds the complete package for a document.
    /// </summary>
    /// <param name="document">The document to serialize.</param>
    /// <returns>The <c>.docx</c> package bytes.</returns>
    public byte[] BuildPackage(WordDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        // Each write gets its own session, so relationship maps, list counters, and part references
        // cannot leak from one document into the next.
        return new DocumentWriteSession(document, Recovery).Build();
    }

    /// <summary>Releases resources used by the writer.</summary>
    public void Dispose() => GC.SuppressFinalize(this);
}
