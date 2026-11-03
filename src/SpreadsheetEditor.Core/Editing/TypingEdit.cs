namespace SpreadsheetEditor.Core.Editing;

/// <summary>
/// Sets one value that the user types, one call per keystroke: an entry's text, an entry's name, a table name. Consecutive edits of the
/// same value (the same <paramref name="key"/>) merge into one step that keeps the first old value; a step that ends back at the old
/// value is empty and the history drops it.
/// </summary>
/// <param name="key">What is being typed in, compared with <see cref="object.Equals(object?)"/>: the object and the property.</param>
/// <param name="set">Stores a value in the model.</param>
/// <param name="oldValue">The value before the first keystroke.</param>
/// <param name="newValue">The value after this keystroke.</param>
/// <param name="change">The change raised after the value was stored, for a do, an undo and a redo alike.</param>
internal sealed class TypingEdit<T>(object key, Action<T> set, T oldValue, T newValue, Func<WorkbookChange> change) : IEdit
{
    private T newValue = newValue;

    public bool IsEmpty => EqualityComparer<T>.Default.Equals(oldValue, newValue);

    public IReadOnlyList<WorkbookChange> Apply()
    {
        set(newValue);
        return [change()];
    }

    public IReadOnlyList<WorkbookChange> Revert()
    {
        set(oldValue);
        return [change()];
    }

    public bool TryMerge(IEdit next)
    {
        if (next is not TypingEdit<T> typing || !Equals(typing.Key, key))
        {
            return false;
        }

        newValue = typing.newValue;
        return true;
    }

    private object Key => key;
}
