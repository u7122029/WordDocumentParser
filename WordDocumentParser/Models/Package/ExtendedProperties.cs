using WordDocumentParser.Core;

namespace WordDocumentParser.Models.Package;

/// <summary>
/// Extended document properties, from <c>docProps/app.xml</c>.
/// </summary>
/// <remarks>
/// <para>
/// The statistics below are whatever the authoring application last wrote. Word recalculates them
/// on its next save; this library preserves them rather than recomputing them.
/// </para>
/// <para>
/// Assignments are tracked (see <see cref="TrackedModel"/>), so the writer can tell a property the
/// caller cleared from one that was simply never set.
/// </para>
/// </remarks>
public class ExtendedProperties : TrackedModel
{
    private string? _template;
    private string? _application;
    private string? _appVersion;
    private string? _company;
    private int? _pages;
    private int? _words;
    private int? _characters;
    private int? _charactersWithSpaces;
    private int? _lines;
    private int? _paragraphs;
    private string? _manager;
    private int? _totalTime;

    /// <summary>Template the document is attached to.</summary>
    public string? Template { get => _template; set => Set(ref _template, value); }

    /// <summary>Application that last saved the document.</summary>
    public string? Application { get => _application; set => Set(ref _application, value); }

    /// <summary>Version of that application.</summary>
    public string? AppVersion { get => _appVersion; set => Set(ref _appVersion, value); }

    /// <summary>Company name.</summary>
    public string? Company { get => _company; set => Set(ref _company, value); }

    /// <summary>Page count as last recorded.</summary>
    public int? Pages { get => _pages; set => Set(ref _pages, value); }

    /// <summary>Word count as last recorded.</summary>
    public int? Words { get => _words; set => Set(ref _words, value); }

    /// <summary>Character count excluding spaces, as last recorded.</summary>
    public int? Characters { get => _characters; set => Set(ref _characters, value); }

    /// <summary>Character count including spaces, as last recorded.</summary>
    public int? CharactersWithSpaces { get => _charactersWithSpaces; set => Set(ref _charactersWithSpaces, value); }

    /// <summary>Line count as last recorded.</summary>
    public int? Lines { get => _lines; set => Set(ref _lines, value); }

    /// <summary>Paragraph count as last recorded.</summary>
    public int? Paragraphs { get => _paragraphs; set => Set(ref _paragraphs, value); }

    /// <summary>Manager name.</summary>
    public string? Manager { get => _manager; set => Set(ref _manager, value); }

    /// <summary>Total editing time in minutes, as last recorded.</summary>
    public int? TotalTime { get => _totalTime; set => Set(ref _totalTime, value); }

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public ExtendedProperties Clone()
    {
        var clone = new ExtendedProperties
        {
            _template = _template,
            _application = _application,
            _appVersion = _appVersion,
            _company = _company,
            _pages = _pages,
            _words = _words,
            _characters = _characters,
            _charactersWithSpaces = _charactersWithSpaces,
            _lines = _lines,
            _paragraphs = _paragraphs,
            _manager = _manager,
            _totalTime = _totalTime
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
