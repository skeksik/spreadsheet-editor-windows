using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>
/// The open document and its edits: every operation on sheets, cards, entries, values, the table name and the picture layout goes through
/// here, so that it can be undone and redone, marks the document modified (docs/spec.md §8.1) and raises <see cref="Changed"/>. Sheets,
/// cards and pictures are identified by their object and entries by their <see cref="Column"/> object, never by name or position (§9.36).
/// An operation that would change nothing returns false and records nothing.
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
    /// Ends the current typing step. Consecutive edits of one entry's text, one entry's name or one table name are one undo step until
    /// another operation, an undo or redo, a save or this call; the app calls it when the field being typed in loses focus. A step whose
    /// typing ends back at the old text is dropped, and does not mark the document modified.
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
    public bool CanRemoveCard(Sheet sheet)
    {
        CheckSheet(sheet);
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
    public bool CanRemoveEntries(Sheet sheet, IReadOnlyCollection<Column> columns)
    {
        CheckSheet(sheet);
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
    /// have the same name. Consecutive renames of one column are one undo step (<see cref="EndTyping"/>), dropped when the name ends back
    /// where it started. False when the name is unchanged.
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

        Do(new TypingEdit<string>(
            (column, nameof(Column.Name)), value => column.Name = value, column.Name, name, () => new EntryRenamed(sheet, column)));
        return true;
    }

    // Values (§2.6, §3.4).

    /// <summary>
    /// Sets the text of <paramref name="card"/>'s entry in <paramref name="column"/>. Consecutive edits of one entry's text are one undo
    /// step (<see cref="EndTyping"/>), dropped when the text ends back where it started. The value as read
    /// (<see cref="Entry.Original"/>) is kept: the writer keeps its type while the text still shows it. False when the text is unchanged.
    /// </summary>
    public bool SetEntryText(Sheet sheet, Card card, Column column, string text)
    {
        var entry = EntryOf(sheet, card, column);
        ArgumentNullException.ThrowIfNull(text);
        if (text == entry.Text)
        {
            return false;
        }

        Do(new TypingEdit<string>(
            (entry, nameof(Entry.Text)), value => entry.Text = value, entry.Text, text, () => new EntryTextChanged(sheet, card, column)));
        return true;
    }

    /// <summary>Sets an entry's fill colour, or with null removes it (<c>No fill</c>, <see cref="CellStyle.NoFill"/>). False when unchanged.</summary>
    public bool SetFill(Sheet sheet, Card card, Column column, RgbColor? fill)
    {
        var entry = EntryOf(sheet, card, column);
        var style = (entry.Style ?? new CellStyle()) with { Fill = fill, NoFill = fill is null };
        return SetStyle(sheet, card, column, entry, style);
    }

    /// <summary>
    /// Sets an entry's text colour, or with null makes it automatic (<see cref="CellStyle.AutomaticTextColor"/>, §9.6). False when
    /// unchanged.
    /// </summary>
    public bool SetTextColor(Sheet sheet, Card card, Column column, RgbColor? color)
    {
        var entry = EntryOf(sheet, card, column);
        var style = (entry.Style ?? new CellStyle()) with { TextColor = color, AutomaticTextColor = color is null };
        return SetStyle(sheet, card, column, entry, style);
    }

    /// <summary>
    /// Appends pictures made from file bytes, kept unchanged (<see cref="PictureFactory.Create"/>; bytes that cannot be read are kept as an
    /// undecodable picture, §9.28), to an entry, in the given order, and marks the column as holding pictures. One undo step.
    /// </summary>
    /// <returns>The new pictures; none, and nothing recorded, when <paramref name="pictures"/> is empty.</returns>
    public IReadOnlyList<EntryImage> AddPictures(Sheet sheet, Card card, Column column, IReadOnlyList<byte[]> pictures)
    {
        var entry = EntryOf(sheet, card, column);
        ArgumentNullException.ThrowIfNull(pictures);
        var added = pictures.Select(PictureFactory.Create).ToArray();
        if (added.Length > 0)
        {
            SetPictures(sheet, card, column, entry, [.. entry.Images, .. added], isImageColumn: true);
        }

        return added;
    }

    /// <summary>Removes one picture, by identity, from an entry.</summary>
    /// <exception cref="ArgumentException">The picture is not in the entry.</exception>
    public void RemovePicture(Sheet sheet, Card card, Column column, EntryImage picture)
    {
        var entry = EntryOf(sheet, card, column);
        var index = IndexOf(entry, picture);
        var images = entry.Images.ToList();
        images.RemoveAt(index);
        SetPictures(sheet, card, column, entry, [.. images], column.IsImageColumn);
    }

    /// <summary>Removes every picture from an entry. False when it has none.</summary>
    public bool RemoveAllPictures(Sheet sheet, Card card, Column column)
    {
        var entry = EntryOf(sheet, card, column);
        if (entry.Images.Count == 0)
        {
            return false;
        }

        SetPictures(sheet, card, column, entry, [], column.IsImageColumn);
        return true;
    }

    /// <summary>
    /// Moves a picture to <paramref name="newIndex"/> among its entry's pictures. False when it is already there. Whether the app offers it
    /// is an open question (§9.45).
    /// </summary>
    /// <exception cref="ArgumentException">The picture is not in the entry.</exception>
    public bool MovePicture(Sheet sheet, Card card, Column column, EntryImage picture, int newIndex)
    {
        var entry = EntryOf(sheet, card, column);
        var oldIndex = IndexOf(entry, picture);
        ArgumentOutOfRangeException.ThrowIfNegative(newIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(newIndex, entry.Images.Count);
        if (newIndex == oldIndex)
        {
            return false;
        }

        var images = entry.Images.ToList();
        images.RemoveAt(oldIndex);
        images.Insert(newIndex, picture);
        SetPictures(sheet, card, column, entry, [.. images], column.IsImageColumn);
        return true;
    }

    // Table name (§2.7).

    /// <summary>
    /// Sets the table name, or with null or empty text clears it (the writer clears its cell, §9.17, or inserts a row for a new name,
    /// §9.18). Consecutive edits are one undo step (<see cref="EndTyping"/>), dropped when the name ends back where it started. False when
    /// unchanged.
    /// </summary>
    public bool SetTableName(Sheet sheet, string? name)
    {
        CheckSheet(sheet);
        if (string.IsNullOrEmpty(name))
        {
            name = null;
        }

        if (name == sheet.TableName)
        {
            return false;
        }

        Do(new TypingEdit<string?>(
            (sheet, nameof(Sheet.TableName)), value => sheet.TableName = value, sheet.TableName, name, () => new TableNameChanged(sheet)));
        return true;
    }

    // Picture layout (§2.8).

    /// <summary>
    /// Sets the width of one picture unit, clamped to <see cref="PictureLimits.MinWidth"/>-<see cref="PictureLimits.MaxWidth"/>. With the
    /// aspect lock on, the height becomes round(width / ratio), clamped. Pictures per row above the new maximum
    /// (<see cref="PictureLimits.MaxImagesPerRow"/>) are reduced to it. One undo step for all three. False when the clamped width is the
    /// current one.
    /// </summary>
    public bool SetPictureWidth(Sheet sheet, int width)
    {
        CheckSheet(sheet);
        var layout = sheet.ImageLayout;
        width = Math.Clamp(width, PictureLimits.MinWidth, PictureLimits.MaxWidth);
        if (width == layout.Width)
        {
            return false;
        }

        var height = sheet.LockedAspectRatio is { } ratio ? ClampHeight(width / ratio) : layout.Height;
        return SetImageLayout(sheet, new ImageLayout(width, height, Math.Min(layout.ImagesPerRow, PictureLimits.MaxImagesPerRow(width))));
    }

    /// <summary>
    /// Sets the height of one picture row, clamped to <see cref="PictureLimits.MinHeight"/>-<see cref="PictureLimits.MaxHeight"/>. With
    /// the aspect lock on, the width becomes round(height × ratio), clamped, and pictures per row are reduced as for
    /// <see cref="SetPictureWidth"/>. False when the clamped height is the current one.
    /// </summary>
    public bool SetPictureHeight(Sheet sheet, int height)
    {
        CheckSheet(sheet);
        var layout = sheet.ImageLayout;
        height = Math.Clamp(height, PictureLimits.MinHeight, PictureLimits.MaxHeight);
        if (height == layout.Height)
        {
            return false;
        }

        var width = sheet.LockedAspectRatio is { } ratio ? ClampWidth(height * ratio) : layout.Width;
        return SetImageLayout(sheet, new ImageLayout(width, height, Math.Min(layout.ImagesPerRow, PictureLimits.MaxImagesPerRow(width))));
    }

    /// <summary>Sets how many picture units fit in a row of a cell, clamped to 1-<see cref="PictureLimits.MaxImagesPerRow"/> of the width. False when unchanged.</summary>
    public bool SetPicturesPerRow(Sheet sheet, int imagesPerRow)
    {
        CheckSheet(sheet);
        var layout = sheet.ImageLayout;
        imagesPerRow = Math.Clamp(imagesPerRow, 1, PictureLimits.MaxImagesPerRow(layout.Width));
        return SetImageLayout(sheet, layout with { ImagesPerRow = imagesPerRow });
    }

    /// <summary>
    /// Turns the aspect lock on, capturing ratio = width / height now, or off. Not an undoable step and does not mark the document
    /// modified (§8.1); the lock is kept in memory on the sheet and never saved (§9.51 is open). Turning it on while it is on keeps the
    /// ratio it captured. Raises <see cref="AspectLockChanged"/>. False when it is already in that state.
    /// </summary>
    public bool SetAspectLock(Sheet sheet, bool locked)
    {
        CheckSheet(sheet);
        if (locked == sheet.LockedAspectRatio.HasValue)
        {
            return false;
        }

        sheet.LockedAspectRatio = locked ? (double)sheet.ImageLayout.Width / sheet.ImageLayout.Height : null;
        Changed?.Invoke(this, new WorkbookChangedEventArgs([new AspectLockChanged(sheet)]));
        return true;
    }

    // Sheets (§2.3).

    /// <summary>
    /// Checks a name for a new or renamed sheet against the worksheet name rules, the workbook's sheets and the worksheets the app does not
    /// show (<see cref="Workbook.SkippedSheetNames"/>), for the dialog's inline reason (§9.39).
    /// </summary>
    /// <param name="name">The proposed name.</param>
    /// <param name="renamed">The sheet being renamed, which may keep its own name with another case; null for a new sheet.</param>
    public SheetNameCheck CheckSheetName(string name, Sheet? renamed = null)
    {
        if (renamed is not null)
        {
            CheckSheet(renamed);
        }

        return SheetNames.Check(name, UsedSheetNames(), renamed?.Name);
    }

    /// <summary>
    /// Appends a new sheet named <paramref name="baseName"/> followed by the smallest unused number (<see cref="SheetNames.NewSheetName"/>),
    /// with the defaults of §2.3: one unnamed entry, one empty card, headers shown, no table name, the default picture layout.
    /// </summary>
    /// <param name="baseName">The localised word for a sheet, supplied by the app.</param>
    public Sheet AddSheet(string baseName)
    {
        ArgumentNullException.ThrowIfNull(baseName);
        var sheet = new Sheet(SheetNames.NewSheetName(baseName, UsedSheetNames()));
        sheet.Columns.Add(new Column(""));
        var card = new Card();
        card.Entries.Add(new Entry());
        sheet.Cards.Add(card);
        Do(InsertSheetEdit(Workbook.Sheets.Count, sheet));
        return sheet;
    }

    /// <summary>
    /// Renames a sheet in place (the writer renames the worksheet, §9.30). A change of case only is allowed. False when the name is
    /// unchanged.
    /// </summary>
    /// <exception cref="SheetNameException">The name breaks a rule (<see cref="CheckSheetName"/>); its <see cref="SheetNameException.Check"/> says which.</exception>
    public bool RenameSheet(Sheet sheet, string name)
    {
        CheckSheet(sheet);
        ArgumentNullException.ThrowIfNull(name);
        if (name == sheet.Name)
        {
            return false;
        }

        var check = CheckSheetName(name, sheet);
        if (!check.IsValid)
        {
            throw new SheetNameException(name, check);
        }

        var oldName = sheet.Name;
        Do(new DelegateEdit(() => RenameSheetTo(sheet, name), () => RenameSheetTo(sheet, oldName)));
        return true;
    }

    /// <summary>Whether <see cref="RemoveSheet"/> may remove a sheet: never the workbook's only sheet.</summary>
    public bool CanRemoveSheet(Sheet sheet)
    {
        CheckSheet(sheet);
        return Workbook.Sheets.Count > 1;
    }

    /// <summary>
    /// Removes a sheet; the writer removes its worksheet. An ordinary undoable step: the modified flag follows the history (whether
    /// removing should ever clear it is open, §9.41).
    /// </summary>
    /// <exception cref="InvalidOperationException">It is the workbook's only sheet.</exception>
    public void RemoveSheet(Sheet sheet)
    {
        if (!CanRemoveSheet(sheet))
        {
            throw new InvalidOperationException("Cannot remove the last sheet.");
        }

        var index = Workbook.Sheets.IndexOf(sheet);
        Do(new DelegateEdit(() => RemoveSheetAt(index, sheet), () => InsertSheetAt(index, sheet)));
    }

    /// <summary>Moves a sheet to <paramref name="newIndex"/>; the saved sheet order follows (§9.31). Marks the document modified. False when it is already there.</summary>
    public bool MoveSheet(Sheet sheet, int newIndex)
    {
        CheckSheet(sheet);
        var oldIndex = Workbook.Sheets.IndexOf(sheet);
        ArgumentOutOfRangeException.ThrowIfNegative(newIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(newIndex, Workbook.Sheets.Count);
        if (newIndex == oldIndex)
        {
            return false;
        }

        Do(new DelegateEdit(() => MoveSheetAt(sheet, oldIndex, newIndex), () => MoveSheetAt(sheet, newIndex, oldIndex)));
        return true;
    }

    private bool SetStyle(Sheet sheet, Card card, Column column, Entry entry, CellStyle style)
    {
        var oldStyle = entry.Style;
        if (style == oldStyle)
        {
            return false;
        }

        Do(new DelegateEdit(
            () =>
            {
                entry.Style = style;
                return [new EntryStyleChanged(sheet, card, column)];
            },
            () =>
            {
                entry.Style = oldStyle;
                return [new EntryStyleChanged(sheet, card, column)];
            }));
        return true;
    }

    private void SetPictures(Sheet sheet, Card card, Column column, Entry entry, EntryImage[] images, bool isImageColumn)
    {
        var oldImages = entry.Images.ToArray();
        var wasImageColumn = column.IsImageColumn;
        Do(new DelegateEdit(
            () =>
            {
                entry.Images.Clear();
                entry.Images.AddRange(images);
                column.IsImageColumn = isImageColumn;
                return [new EntryPicturesChanged(sheet, card, column)];
            },
            () =>
            {
                entry.Images.Clear();
                entry.Images.AddRange(oldImages);
                column.IsImageColumn = wasImageColumn;
                return [new EntryPicturesChanged(sheet, card, column)];
            }));
    }

    private bool SetImageLayout(Sheet sheet, ImageLayout layout)
    {
        var oldLayout = sheet.ImageLayout;
        if (layout == oldLayout)
        {
            return false;
        }

        Do(new DelegateEdit(
            () =>
            {
                sheet.ImageLayout = layout;
                return [new PictureLayoutChanged(sheet)];
            },
            () =>
            {
                sheet.ImageLayout = oldLayout;
                return [new PictureLayoutChanged(sheet)];
            }));
        return true;
    }

    // §2.8's "round" is taken as rounding half away from zero; which rounding the original app used is not verified.
    private static int ClampWidth(double width) =>
        (int)Math.Clamp(Math.Round(width, MidpointRounding.AwayFromZero), PictureLimits.MinWidth, PictureLimits.MaxWidth);

    private static int ClampHeight(double height) =>
        (int)Math.Clamp(Math.Round(height, MidpointRounding.AwayFromZero), PictureLimits.MinHeight, PictureLimits.MaxHeight);

    private IEnumerable<string> UsedSheetNames() => Workbook.Sheets.Select(sheet => sheet.Name).Concat(Workbook.SkippedSheetNames);

    private DelegateEdit InsertSheetEdit(int index, Sheet sheet) => new(() => InsertSheetAt(index, sheet), () => RemoveSheetAt(index, sheet));

    private WorkbookChange[] InsertSheetAt(int index, Sheet sheet)
    {
        Workbook.Sheets.Insert(index, sheet);
        return [new SheetInserted(sheet, index)];
    }

    private WorkbookChange[] RemoveSheetAt(int index, Sheet sheet)
    {
        Workbook.Sheets.RemoveAt(index);
        return [new SheetRemoved(sheet, index)];
    }

    private static WorkbookChange[] RenameSheetTo(Sheet sheet, string name)
    {
        sheet.Name = name;
        return [new SheetRenamed(sheet)];
    }

    private WorkbookChange[] MoveSheetAt(Sheet sheet, int from, int to)
    {
        Workbook.Sheets.RemoveAt(from);
        Workbook.Sheets.Insert(to, sheet);
        return [new SheetMoved(sheet, from, to)];
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

    private Entry EntryOf(Sheet sheet, Card card, Column column)
    {
        IndexOf(sheet, card);
        return card.Entries[IndexOf(sheet, column)];
    }

    private static int IndexOf(Entry entry, EntryImage picture)
    {
        ArgumentNullException.ThrowIfNull(picture);
        var index = entry.Images.IndexOf(picture);
        return index >= 0 ? index : throw new ArgumentException("The picture is not in this entry.", nameof(picture));
    }

    private int IndexOf(Sheet sheet, Column column)
    {
        CheckSheet(sheet);
        ArgumentNullException.ThrowIfNull(column);
        var index = sheet.Columns.IndexOf(column);
        return index >= 0 ? index : throw new ArgumentException("The entry is not in this sheet.", nameof(column));
    }
}
