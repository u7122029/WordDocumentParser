namespace WordDocumentParser.Core;

/// <summary>
/// Interface for writing document trees to Word documents.
/// </summary>
public interface IDocumentWriter : IDisposable
{
    /// <summary>
    /// Writes a document tree to a file.
    /// </summary>
    /// <param name="root">Root node of the document tree</param>
    /// <param name="filePath">Path to write the .docx file</param>
    void WriteToFile(DocumentNode root, string filePath);

    /// <summary>
    /// Writes a document tree to a stream.
    /// </summary>
    /// <param name="root">Root node of the document tree</param>
    /// <param name="stream">Stream to write the .docx data</param>
    void WriteToStream(DocumentNode root, Stream stream);
}
