namespace WordDocumentParser.Core;

/// <summary>
/// Interface for parsing Word documents into a document tree structure.
/// </summary>
public interface IDocumentParser : IDisposable
{
    /// <summary>
    /// Parses a Word document from a file path.
    /// </summary>
    /// <param name="filePath">Path to the .docx file</param>
    /// <returns>Root node of the document tree</returns>
    DocumentNode ParseFromFile(string filePath);

    /// <summary>
    /// Parses a Word document from a stream.
    /// </summary>
    /// <param name="stream">Stream containing the .docx data</param>
    /// <param name="documentName">Optional name for the document</param>
    /// <returns>Root node of the document tree</returns>
    DocumentNode ParseFromStream(Stream stream, string documentName = "Document");
}
