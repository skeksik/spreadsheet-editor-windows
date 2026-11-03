using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>
/// The open document and its edits: every operation on cards and entries goes through here, so that it can be undone and redone, marks
/// the document modified (docs/spec.md §8.1) and raises <see cref="Changed"/>. Cards are identified by the <see cref="Card"/> object and
/// entries by their <see cref="Column"/> object, never by name or position (§9.36).
/// </summary>
/// <remarks>
/// Not thread-safe: the document is edited on the UI thread only, and every event is raised on the calling thread (ADR 0002).
/// </remarks>
public sealed class WorkbookEditor
{
    private readonly UndoHistory history = new();

    public WorkbookEditor(Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        Workbook = workbook;
    }

    /// <summary>Raised after an operation, an undo or a redo changed the model.</summary>
    public event EventHandler<WorkbookChangedEventArgs>? Changed;

    /// <summary>Raised after <see cref="CanUndo"/>, <see cref="CanRedo"/> or <see cref="IsModified"/> may have changed.</summary>
    public event EventHandler? StateChanged;

    public Workbook Workbook { get; private set; }

    public bool CanUndo => history.CanUndo;

    public bool CanRedo => history.CanRedo;

    /// <summary>
    /// True when the document differs from what was last loaded or saved: set by every operation, cleared by <see cref="Load"/>, by
    /// <see cref="MarkSaved"/> and by undoing or redoing back to the saved state (§8.1).
    /// </summary>
    public bool IsModified => history.IsModified;

    /// <summary>Replaces the document with a newly read or new workbook: no undo or redo, unmodified. Raises only <see cref="StateChanged"/>.</summary>
    public void Load(Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        Workbook = workbook;
        history.Clear();
        RaiseStateChanged();
    }

    /// <summary>Records that the document was saved in its current state; the app calls it after a successful save.</summary>
    public void MarkSaved()
    {
        history.MarkSaved();
        RaiseStateChanged();
    }

    /// <summary>
    /// Ends the current typing step. Consecutive renames of one entry are one undo step until another operation, an undo or redo, a save
    /// or this call; the app calls it when the field being typed in loses focus.
    /// </summary>
    public void EndTyping() => history.BreakMerge();

    public void Undo()
    {
        var edit = history.PopUndo() ?? throw new InvalidOperationException("There is nothing to undo.");
        Raise(edit.Revert());
    }

    public void Redo()
    {
        var edit = history.PopRedo() ?? throw new InvalidOperationException("There is nothing to redo.");
        Raise(edit.Apply());
    }

    // Cards (§2.4).

    /// <summary>Appends an empty card: one empty entry per column, row span 1.</summary>
    public Card AddCard(Sheet sheet)
    {
        CheckSheet(sheet);
        return InsertNewCard(sheet, sheet.Cards.Count);
    }

    /// <summary>Inserts an empty card directly above <paramref name="card"/>.</summary>
    public Card InsertCardAbove(Sheet sheet, Card card) => InsertNewCard(sheet, IndexOf(sheet, card));

    /// <summary>Inserts an empty card directly below <paramref name="card"/>.</summary>
    public Card InsertCardBelow(Sheet sheet, Card card) => InsertNewCard(sheet, IndexOf(sheet, card) + 1);

    /// <summary>
    /// Inserts a copy of <paramref name="card"/> directly below it: its text, typed values, styles, pictures and row span (§9.22). The copy
    /// is a new card, not read from the file.
    /// </summary>
    public Card DuplicateCard(Sheet sheet, Card card)
    {
        var index = IndexOf(sheet, card);
        var copy = new Card { SourceRowSpan = card.SourceRowSpan };
        foreach (var entry in card.Entries)
        {
            var entryCopy = new Entry { Text = entry.Text, Original = entry.Original, Style = entry.Style };
            entryCopy.Images.AddRange(entry.Images);
            copy.Entries.Add(entryCopy);
        }

        Do(InsertCardEdit(sheet, index + 1, copy));
        return copy;
    }

    /// <summary>Whether <see cref="RemoveCard"/> may remove a card of <paramref name="sheet"/>: never its only card (§2.4).</summary>
    public static bool CanRemoveCard(Sheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        return sheet.Cards.Count > 1;
    }

    /// <summary>Removes a card. Its rows go with it when the file is saved.</summary>
    /// <exception cref="InvalidOperationException">It is the sheet's only card (§2.4).</exception>
    public void RemoveCard(Sheet sheet, Card card)
    {
        var index = IndexOf(sheet, card);
        if (!CanRemoveCard(sheet))
        {
            throw new InvalidOperationException("Cannot remove the last card.");
        }

        Do(new DelegateEdit(
            () =>
            {
                sheet.Cards.RemoveAt(index);
                return [new CardRemoved(sheet, index, card)];
            },
            () =>
            {
                sheet.Cards.Insert(index, card);
                return [new CardInserted(sheet, index, card)];
            }));
    }

    /// <summary>Moves a card to <paramref name="newIndex"/> in the sheet's cards; its row span moves with it. False when it is already there.</summary>
    public bool MoveCard(Sheet sheet, Card card, int newIndex)
    {
        var oldIndex = IndexOf(sheet, card);
        ArgumentOutOfRangeException.ThrowIfNegative(newIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(newIndex, sheet.Cards.Count);
        if (newIndex == oldIndex)
        {
            return false;
        }

        Do(new DelegateEdit(() => MoveCardAt(sheet, card, oldIndex, newIndex), () => MoveCardAt(sheet, card, newIndex, oldIndex)));
        return true;
    }

    public bool CanMoveCardUp(Sheet sheet, Card card) => IndexOf(sheet, card) > 0;

    public bool CanMoveCardDown(Sheet sheet, Card card) => IndexOf(sheet, card) < sheet.Cards.Count - 1;

    /// <summary>Swaps a card with the one above it. On the first card nothing happens and nothing is recorded; returns false.</summary>
    public bool MoveCardUp(Sheet sheet, Card card)
    {
        var index = IndexOf(sheet, card);
        return index > 0 && MoveCard(sheet, card, index - 1);
    }

    /// <summary>Swaps a card with the one below it. On the last card nothing happens and nothing is recorded; returns false.</summary>
    public bool MoveCardDown(Sheet sheet, Card card)
    {
        var index = IndexOf(sheet, card);
        return index < sheet.Cards.Count - 1 && MoveCard(sheet, card, index + 1);
    }

    // Entries (§2.5).

    /// <summary>Appends an unnamed column (empty name, §9.38) and an empty entry to every card. Row spans are unchanged.</summary>
    public Column AddEntry(Sheet sheet)
    {
        CheckSheet(sheet);
        var column = new Column("");
        var entries = sheet.Cards.Select(_ => new Entry()).ToArray();
        Do(new DelegateEdit(
            () =>
            {
                sheet.Columns.Add(column);
                for (var i = 0; i < sheet.Cards.Count; i++)
                {
                    sheet.Cards[i].Entries.Add(entries[i]);
                }

                return [new EntriesChanged(sheet)];
            },
            () =>
            {
                sheet.Columns.RemoveAt(sheet.Columns.Count - 1);
                foreach (var card in sheet.Cards)
                {
                    card.Entries.RemoveAt(card.Entries.Count - 1);
                }

                return [new EntriesChanged(sheet)];
            }));
        return column;
    }

    /// <summary>Whether <see cref="RemoveEntries"/> may remove these columns: at least one, and never every column (§9.33).</summary>
    public static bool CanRemoveEntries(Sheet sheet, IReadOnlyCollection<Column> columns)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(columns);
        var removed = sheet.Columns.Count(columns.Contains);
        return removed > 0 && removed < sheet.Columns.Count;
    }

    /// <summary>Removes the given columns, by identity, and their entry in every card. Row spans are unchanged.</summary>
    /// <exception cref="ArgumentException">A column is not in the sheet, or none is given.</exception>
    /// <exception cref="InvalidOperationException">It would remove every column (§9.33).</exception>
    public void RemoveEntries(Sheet sheet, IReadOnlyCollection<Column> columns)
    {
        CheckSheet(sheet);
        ArgumentNullException.ThrowIfNull(columns);
        if (columns.Count == 0)
        {
            throw new ArgumentException("No entry to remove was given.", nameof(columns));
        }

        var indexes = columns.Select(column => IndexOf(sheet, column)).Distinct().Order().ToArray();
        if (indexes.Length == sheet.Columns.Count)
        {
            throw new InvalidOperationException("Cannot remove the last entry.");
        }

        var removedColumns = indexes.Select(i => sheet.Columns[i]).ToArray();
        var removedEntries = sheet.Cards.Select(card => indexes.Select(i => card.Entries[i]).ToArray()).ToArray();
        Do(new DelegateEdit(
            () =>
            {
                for (var k = indexes.Length - 1; k >= 0; k--)
                {
                    sheet.Columns.RemoveAt(indexes[k]);
                    foreach (var card in sheet.Cards)
                    {
                        card.Entries.RemoveAt(indexes[k]);
                    }
                }

                return [new EntriesChanged(sheet)];
            },
            () =>
            {
                for (var k = 0; k < indexes.Length; k++)
                {
                    sheet.Columns.Insert(indexes[k], removedColumns[k]);
                    for (var c = 0; c < sheet.Cards.Count; c++)
                    {
                        sheet.Cards[c].Entries.Insert(indexes[k], removedEntries[c][k]);
                    }
                }

                return [new EntriesChanged(sheet)];
            }));
    }

    /// <summary>
    /// Puts the columns in the given order, and every card's entries with them. <paramref name="order"/> holds each of the sheet's columns
    /// once. Row spans are unchanged (§9.22). False when the order is the current one.
    /// </summary>
    /// <exception cref="ArgumentException"><paramref name="order"/> is not an order of the sheet's columns.</exception>
    public bool ReorderEntries(Sheet sheet, IReadOnlyList<Column> order)
    {
        CheckSheet(sheet);
        ArgumentNullException.ThrowIfNull(order);
        if (order.Count != sheet.Columns.Count || order.Distinct().Count() != order.Count || !order.All(sheet.Columns.Contains))
        {
            throw new ArgumentException("The new order must hold each of the sheet's entries once.", nameof(order));
        }

        if (order.SequenceEqual(sheet.Columns))
        {
            return false;
        }

        var oldOrder = sheet.Columns.ToArray();
        var newOrder = order.ToArray();
        Do(new DelegateEdit(() => SetColumnOrder(sheet, newOrder), () => SetColumnOrder(sheet, oldOrder)));
        return true;
    }

    /// <summary>
    /// Renames a column; every card shows the new name. Any text is a name, and only the empty name is unnamed (§9.38); two columns may
    /// have the same name. Consecutive renames of one column are one undo step (<see cref="EndTyping"/>). False when the name is unchanged.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// The sheet's headers are hidden (§3.3): its entries have no names to edit. Whether naming them should turn headers on is an open
    /// question (§9.37); until it is answered, the spec's current behaviour holds.
    /// </exception>
    public bool RenameEntry(Sheet sheet, Column column, string name)
    {
        IndexOf(sheet, column);
        ArgumentNullException.ThrowIfNull(name);
        if (sheet.HeaderRow is null)
        {
            throw new InvalidOperationException("Entries on a sheet with hidden headers have no names.");
        }

        if (name == column.Name)
        {
            return false;
        }

        Do(new RenameEntryEdit(sheet, column, column.Name, name));
        return true;
    }

    private Card InsertNewCard(Sheet sheet, int index)
    {
        var card = new Card();
        card.Entries.AddRange(sheet.Columns.Select(_ => new Entry()));
        Do(InsertCardEdit(sheet, index, card));
        return card;
    }

    private static DelegateEdit InsertCardEdit(Sheet sheet, int index, Card card) => new(
        () =>
        {
            sheet.Cards.Insert(index, card);
            return [new CardInserted(sheet, index, card)];
        },
        () =>
        {
            sheet.Cards.RemoveAt(index);
            return [new CardRemoved(sheet, index, card)];
        });

    private static WorkbookChange[] MoveCardAt(Sheet sheet, Card card, int from, int to)
    {
        sheet.Cards.RemoveAt(from);
        sheet.Cards.Insert(to, card);
        return [new CardMoved(sheet, card, from, to)];
    }

    private static WorkbookChange[] SetColumnOrder(Sheet sheet, Column[] order)
    {
        var oldIndex = order.Select(column => sheet.Columns.IndexOf(column)).ToArray();
        foreach (var card in sheet.Cards)
        {
            var entries = oldIndex.Select(i => card.Entries[i]).ToArray();
            card.Entries.Clear();
            card.Entries.AddRange(entries);
        }

        sheet.Columns.Clear();
        sheet.Columns.AddRange(order);
        return [new EntriesChanged(sheet)];
    }

    private void Do(IEdit edit)
    {
        var changes = edit.Apply();
        history.Record(edit);
        Raise(changes);
    }

    private void Raise(IReadOnlyList<WorkbookChange> changes)
    {
        Changed?.Invoke(this, new WorkbookChangedEventArgs(changes));
        RaiseStateChanged();
    }

    private void RaiseStateChanged() => StateChanged?.Invoke(this, EventArgs.Empty);

    private void CheckSheet(Sheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (!Workbook.Sheets.Contains(sheet))
        {
            throw new ArgumentException("The sheet is not in this workbook.", nameof(sheet));
        }
    }

    private int IndexOf(Sheet sheet, Card card)
    {
        CheckSheet(sheet);
        ArgumentNullException.ThrowIfNull(card);
        var index = sheet.Cards.IndexOf(card);
        return index >= 0 ? index : throw new ArgumentException("The card is not in this sheet.", nameof(card));
    }

    private int IndexOf(Sheet sheet, Column column)
    {
        CheckSheet(sheet);
        ArgumentNullException.ThrowIfNull(column);
        var index = sheet.Columns.IndexOf(column);
        return index >= 0 ? index : throw new ArgumentException("The entry is not in this sheet.", nameof(column));
    }
}
