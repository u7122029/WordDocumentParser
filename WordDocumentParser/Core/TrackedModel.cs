using System.Runtime.CompilerServices;

namespace WordDocumentParser.Core;

/// <summary>
/// Base class for models that record which properties have been assigned since the model
/// was last accepted as the baseline.
/// </summary>
/// <remarks>
/// The parser populates a model from the source document and then calls <see cref="AcceptChanges"/>,
/// so a freshly parsed model reports no changes. Every later assignment through the public API is
/// recorded. The writer consults that record to decide what to rewrite: properties that were never
/// assigned are left as the original XML had them. Without this record the writer cannot tell a
/// value it merely parsed from one the caller actually set, and has to rewrite everything.
/// </remarks>
public abstract class TrackedModel
{
    private HashSet<string>? _changed;

    /// <summary>
    /// Names of the properties assigned since the last <see cref="AcceptChanges"/> call.
    /// </summary>
    public IReadOnlyCollection<string> ChangedProperties => _changed ?? (IReadOnlyCollection<string>)[];

    /// <summary>
    /// True when at least one tracked property has been assigned since the last <see cref="AcceptChanges"/> call.
    /// Derived types that own nested models override this to include them.
    /// </summary>
    public virtual bool HasChanges => HasOwnChanges;

    /// <summary>
    /// True when a property of this instance itself has been assigned, ignoring any nested models.
    /// </summary>
    protected bool HasOwnChanges => _changed is { Count: > 0 };

    /// <summary>
    /// Returns true when the named property has been assigned since the last <see cref="AcceptChanges"/> call.
    /// </summary>
    /// <param name="propertyName">The property name to test.</param>
    public bool IsChanged(string propertyName) => _changed?.Contains(propertyName) is true;

    /// <summary>
    /// Returns true when any of the named properties has been assigned since the last
    /// <see cref="AcceptChanges"/> call.
    /// </summary>
    /// <param name="propertyNames">The property names to test.</param>
    public bool IsAnyChanged(params string[] propertyNames)
    {
        if (_changed is not { Count: > 0 }) return false;
        foreach (var name in propertyNames)
        {
            if (_changed.Contains(name)) return true;
        }
        return false;
    }

    /// <summary>
    /// Clears the change record, treating the current values as the unmodified baseline.
    /// Called by the parser once a model has been populated from the source document.
    /// </summary>
    public void AcceptChanges() => _changed?.Clear();

    /// <summary>
    /// Records the named property as explicitly assigned.
    /// Use this when a value is mutated in place rather than through its property setter.
    /// </summary>
    /// <param name="propertyName">The property name to record.</param>
    public void MarkChanged(string propertyName) =>
        (_changed ??= new HashSet<string>(StringComparer.Ordinal)).Add(propertyName);

    /// <summary>
    /// Copies the change record of another model of the same shape onto this one, so that a clone
    /// carries the pending edits of its source rather than presenting them as a clean baseline.
    /// </summary>
    /// <param name="other">The model whose change record should be copied.</param>
    protected void CopyChangesFrom(TrackedModel other)
    {
        if (other._changed is not { Count: > 0 }) return;
        foreach (var name in other._changed)
        {
            MarkChanged(name);
        }
    }

    /// <summary>
    /// Assigns a backing field and records the change when the value actually differs.
    /// </summary>
    /// <typeparam name="T">The property type.</typeparam>
    /// <param name="field">The backing field to assign.</param>
    /// <param name="value">The new value.</param>
    /// <param name="propertyName">Supplied automatically by the compiler.</param>
    protected void Set<T>(ref T field, T value, [CallerMemberName] string propertyName = "")
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return;
        field = value;
        MarkChanged(propertyName);
    }
}
