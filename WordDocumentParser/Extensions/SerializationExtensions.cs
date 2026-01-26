namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for saving and exporting documents.
/// </summary>
public static class SerializationExtensions
{
    /// <summary>
    /// Saves the document to a Word document file (.docx).
    /// </summary>
    /// <param name="document">The Word document to save</param>
    /// <param name="filePath">Path where the document will be saved</param>
    public static void SaveToFile(this WordDocument document, string filePath)
    {
        using var writer = new WordDocumentTreeWriter();
        writer.WriteToFile(document, filePath);
    }

    /// <summary>
    /// Saves the document to a stream as a Word document (.docx).
    /// </summary>
    /// <param name="document">The Word document to save</param>
    /// <param name="stream">Stream to write to</param>
    public static void SaveToStream(this WordDocument document, Stream stream)
    {
        using var writer = new WordDocumentTreeWriter();
        writer.WriteToStream(document, stream);
    }

    /// <summary>
    /// Saves the document to a byte array as a Word document (.docx).
    /// </summary>
    /// <param name="document">The Word document to save</param>
    /// <returns>Document as a byte array</returns>
    public static byte[] ToDocxBytes(this WordDocument document)
    {
        using var stream = new MemoryStream();
        document.SaveToStream(stream);
        return stream.ToArray();
    }
}
