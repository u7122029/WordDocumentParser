namespace WordDocumentParser.Core;

/// <summary>
/// Interface for writing document trees to Word documents.
/// </summary>
public interface IDocumentWriter : IDisposable
{
    /// <summary>
    /// Writes a document to a file.
    /// </summary>
    /// <param name="document">The Word document to write</param>
    /// <param name="filePath">Path to write the .docx file</param>
    void WriteToFile(WordDocument document, string filePath);

    /// <summary>
    /// Writes a document to a stream.
    /// </summary>
    /// <param name="document">The Word document to write</param>
    /// <param name="stream">Stream to write the .docx data</param>
    void WriteToStream(WordDocument document, Stream stream);
}
