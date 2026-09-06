using WordDocumentParser.Core;

namespace WordDocumentParser.Models.Package;

/// <summary>
/// Core document properties, from <c>docProps/core.xml</c>.
/// </summary>
/// <remarks>
/// Assignments are tracked (see <see cref="TrackedModel"/>), so the writer can tell a property the
/// caller cleared from one that was simply never set. Without the distinction a removal had no
/// effect on the saved package, which kept the original value.
/// </remarks>
public class CoreProperties : TrackedModel
{
    private string? _title;
    private string? _subject;
    private string? _creator;
    private string? _keywords;
    private string? _description;
    private string? _lastModifiedBy;
    private string? _revision;
    private string? _created;
    private string? _modified;
    private string? _category;
    private string? _contentStatus;

    /// <summary>Document title.</summary>
    public string? Title { get => _title; set => Set(ref _title, value); }

    /// <summary>Document subject.</summary>
    public string? Subject { get => _subject; set => Set(ref _subject, value); }

    /// <summary>Author who created the document.</summary>
    public string? Creator { get => _creator; set => Set(ref _creator, value); }

    /// <summary>Keywords, as a single delimited string.</summary>
    public string? Keywords { get => _keywords; set => Set(ref _keywords, value); }

    /// <summary>Free-text description, shown by Word as "Comments".</summary>
    public string? Description { get => _description; set => Set(ref _description, value); }

    /// <summary>Author of the most recent save.</summary>
    public string? LastModifiedBy { get => _lastModifiedBy; set => Set(ref _lastModifiedBy, value); }

    /// <summary>Revision number, as a string.</summary>
    public string? Revision { get => _revision; set => Set(ref _revision, value); }

    /// <summary>Creation timestamp, in ISO 8601 round-trip format.</summary>
    public string? Created { get => _created; set => Set(ref _created, value); }

    /// <summary>Last-modified timestamp, in ISO 8601 round-trip format.</summary>
    public string? Modified { get => _modified; set => Set(ref _modified, value); }

    /// <summary>Document category.</summary>
    public string? Category { get => _category; set => Set(ref _category, value); }

    /// <summary>Content status, for example <c>"Draft"</c> or <c>"Final"</c>.</summary>
    public string? ContentStatus { get => _contentStatus; set => Set(ref _contentStatus, value); }

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public CoreProperties Clone()
    {
        var clone = new CoreProperties
        {
            _title = _title,
            _subject = _subject,
            _creator = _creator,
            _keywords = _keywords,
            _description = _description,
            _lastModifiedBy = _lastModifiedBy,
            _revision = _revision,
            _created = _created,
            _modified = _modified,
            _category = _category,
            _contentStatus = _contentStatus
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
