namespace WordDocumentParser.Models.ContentControls;

/// <summary>
/// Which part of the package a document property lives in.
/// </summary>
public enum DocumentPropertyType
{
    /// <summary>Core property from <c>docProps/core.xml</c>: title, subject, author, and so on.</summary>
    Core,

    /// <summary>Extended property from <c>docProps/app.xml</c>: company, manager, statistics.</summary>
    Extended,

    /// <summary>User-defined property from <c>docProps/custom.xml</c>.</summary>
    Custom
}
