using SpreadsheetEditor.Core.Formulas;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Where each worksheet row of a sheet read from <c>.xlsx</c> goes when its table is written (docs/spec.md §5.3, §9.18, §9.20, §9.22): the
/// cards in the model's order, each with its own whole rows, new cards in new rows, removed cards' rows deleted, and everything below the
/// table moved so that it keeps its place relative to the table's end.
/// </summary>
/// <remarks>
/// <para>The change is also expressed as a sequence of <see cref="RowShift"/>s, which every reference to the sheet's rows goes through
/// (formulas, defined names, ranges): first the removed cards' rows are deleted, bottom up, then the new rows are inserted, top down, at
/// their final rows (the table name row first). Reordering the cards that stay is not in that sequence: a reference to a row of the table
/// keeps pointing at that position, as after a sort in a spreadsheet program, whichever card ends up there (manager, 2026-11-03). Only what
/// belongs to a row (its cells, height and style, the merges, hyperlinks, comments and pictures inside one card's rows) moves with its
/// card; <see cref="Map"/> gives that move and <see cref="Positional"/> the position the shifts alone give.</para>
/// <para>A card read from the file keeps its rows and gets more when its <see cref="Card.SourceRowSpan"/> asks for more; any other card
/// (new or duplicated, §9.22) gets <see cref="Card.SourceRowSpan"/> new rows.</para>
/// </remarks>
internal sealed class RowPlan
{
    private readonly Dictionary<int, int> kept = [];
    private readonly Dictionary<int, int> keptBack = [];
    private readonly Dictionary<int, int> positional = [];
    private readonly Dictionary<int, (int First, int Last)> slotOf = [];
    private readonly HashSet<int> removed = [];
    private int dataLastRow;
    private bool insertsName;

    /// <summary>The row shifts in the order they apply: deletions (bottom up), then insertions (top down).</summary>
    public List<RowShift> Shifts { get; } = [];

    /// <summary>One entry per model card, in order.</summary>
    public List<CardRows> Cards { get; } = [];

    /// <summary>The removed cards' rows as read, as (first, last) pairs.</summary>
    public List<(int First, int Last)> RemovedRows { get; } = [];

    /// <summary>The row the table name is written in, after the change; null when the table has no name row and gets none.</summary>
    public int? NameRow { get; private set; }

    /// <summary>The header row after the change, or null when the headers are hidden.</summary>
    public int? HeaderRow { get; private set; }

    /// <summary>The table's first row after the change: its name row, else its header row, else its first card.</summary>
    public int FirstRow { get; private set; }

    /// <summary>The first card's row after the change.</summary>
    public int FirstDataRow { get; private set; }

    /// <summary>The table's last row after the change (its last card's last row, or the header row when it has no card row).</summary>
    public int LastRow { get; private set; }

    /// <summary>True when any row moves, is inserted or is deleted.</summary>
    public bool ChangesRows { get; private set; }

    /// <param name="sheet">The sheet as it is written.</param>
    /// <param name="origin">Where its table was as read.</param>
    /// <param name="slots">The cards' rows as read, as (first row, row count) pairs.</param>
    /// <param name="pictureRows">
    /// The worksheet rows a card needs for its pictures, given the first row as read of the rows it keeps (null for a card in new rows);
    /// 0 when its pictures need none. A card gets them when it has fewer (§5.3 step 3); rows are never deleted for fewer pictures.
    /// </param>
    public static RowPlan Create(
        Sheet sheet, SheetSource origin, List<(int FirstRow, int RowSpan)> slots, Func<Card, int?, int>? pictureRows = null)
    {
        var firstTableRow = sheet.HeaderRow ?? sheet.FirstDataRow;
        var hasNameRow = origin.FirstRow < firstTableRow;
        var insertName = !hasNameRow && !string.IsNullOrEmpty(sheet.TableName);
        var plan = new RowPlan();
        var nameShift = insertName ? 1 : 0;
        plan.NameRow = insertName ? firstTableRow : hasNameRow ? origin.FirstRow : null;
        plan.HeaderRow = sheet.HeaderRow + nameShift;
        plan.FirstDataRow = sheet.FirstDataRow + nameShift;
        plan.FirstRow = plan.NameRow ?? plan.HeaderRow ?? plan.FirstDataRow;

        var slotIndex = new Dictionary<int, int>();
        for (var i = 0; i < slots.Count; i++)
        {
            slotIndex[slots[i].FirstRow] = i;
            for (var row = slots[i].FirstRow; row < slots[i].FirstRow + slots[i].RowSpan; row++)
            {
                plan.slotOf[row] = (slots[i].FirstRow, slots[i].FirstRow + slots[i].RowSpan - 1);
            }
        }

        // A table with a header and no card was read with one empty card that has no row in the file; while it stays empty it gets none.
        var placeholder = slots.Count == 0 && sheet.Cards is [{ SourceFirstRow: null } only]
            && only.Entries.TrueForAll(e => e.Text.Length == 0 && e.Images.Count == 0);

        // Which slot each card takes: the first card that names a slot as its own takes it; any other card gets new rows.
        var claimed = new bool[slots.Count];
        var layout = new List<(int? Slot, int Rows)>();
        foreach (var card in placeholder ? [] : sheet.Cards)
        {
            if (card.SourceFirstRow is { } first && slotIndex.TryGetValue(first, out var index) && !claimed[index])
            {
                claimed[index] = true;
                var need = pictureRows?.Invoke(card, slots[index].FirstRow) ?? 0;
                layout.Add((index, Math.Max(Math.Max(slots[index].RowSpan, card.SourceRowSpan), need)));
            }
            else
            {
                var need = pictureRows?.Invoke(card, null) ?? 0;
                layout.Add((null, Math.Max(Math.Max(1, card.SourceRowSpan), need)));
            }
        }

        // Deletions, bottom up in the rows as read; adjacent removed cards are deleted at once.
        var deletions = new List<(int First, int Count)>();
        for (var i = 0; i < slots.Count; i++)
        {
            if (claimed[i])
            {
                continue;
            }

            var (first, span) = slots[i];
            plan.RemovedRows.Add((first, first + span - 1));
            for (var row = first; row < first + span; row++)
            {
                plan.removed.Add(row);
            }

            if (deletions.Count > 0 && deletions[^1].First + deletions[^1].Count == first)
            {
                deletions[^1] = (deletions[^1].First, deletions[^1].Count + span);
            }
            else
            {
                deletions.Add((first, span));
            }
        }

        for (var i = deletions.Count - 1; i >= 0; i--)
        {
            plan.Shifts.Add(RowShift.Delete(sheet.Name, deletions[i].First, deletions[i].Count));
        }

        if (insertName)
        {
            plan.Shifts.Add(RowShift.Insert(sheet.Name, firstTableRow, 1));
        }

        // The final layout: each card's rows in order. A kept card's own rows are "kept positions"; its extra rows and new cards' rows are
        // inserted. Positionally the kept rows fill the kept positions in the order they were read.
        var keptPositions = new List<int>();
        var inserted = new List<int>();
        var cursor = plan.FirstDataRow;
        foreach (var (slot, rows) in layout)
        {
            var own = slot is { } s ? slots[s].RowSpan : 0;
            if (slot is { } index)
            {
                for (var k = 0; k < own; k++)
                {
                    plan.kept[slots[index].FirstRow + k] = cursor + k;
                    plan.keptBack[cursor + k] = slots[index].FirstRow + k;
                    keptPositions.Add(cursor + k);
                }
            }

            for (var k = own; k < rows; k++)
            {
                inserted.Add(cursor + k);
            }

            plan.Cards.Add(new CardRows(cursor, rows, slot is { } i ? slots[i].FirstRow : null, own));
            cursor += rows;
        }

        var keptOld = plan.kept.Keys.Order().ToList();
        for (var j = 0; j < keptOld.Count; j++)
        {
            plan.positional[keptOld[j]] = keptPositions[j];
        }

        for (var i = 0; i < inserted.Count;)
        {
            var start = inserted[i];
            var count = 1;
            while (i + count < inserted.Count && inserted[i + count] == start + count)
            {
                count++;
            }

            plan.Shifts.Add(RowShift.Insert(sheet.Name, start, count));
            i += count;
        }

        plan.dataLastRow = cursor - 1;
        plan.insertsName = insertName;
        plan.LastRow = Math.Max(cursor - 1, plan.HeaderRow ?? plan.FirstRow);
        plan.ChangesRows = plan.Shifts.Count > 0 || plan.kept.Any(pair => pair.Key != pair.Value);
        return plan;
    }

    /// <summary>The row a row as read moves to, or null when it is deleted (a removed card's row).</summary>
    public int? Map(int row)
    {
        if (kept.TryGetValue(row, out var moved))
        {
            return moved;
        }

        return removed.Contains(row) ? null : ShiftPoint(row);
    }

    /// <summary>The row the shifts alone move a row as read to (references are moved so), or null when it is deleted.</summary>
    public int? Positional(int row)
    {
        if (positional.TryGetValue(row, out var moved))
        {
            return moved;
        }

        return removed.Contains(row) ? null : ShiftPoint(row);
    }

    /// <summary>
    /// The row as read whose content is at <paramref name="row"/> after the change, or null for a new row (an inserted row, or the rows
    /// of a new card).
    /// </summary>
    public int? OldRow(int row)
    {
        if (keptBack.TryGetValue(row, out var old))
        {
            return old;
        }

        if ((row >= FirstDataRow && row <= dataLastRow) || (insertsName && row == NameRow))
        {
            // Inside the table's data rows every row that is not a kept card's is new, and so is an inserted name row.
            return null;
        }

        int? current = row;
        for (var i = Shifts.Count - 1; i >= 0 && current is { } r; i--)
        {
            var shift = Shifts[i];
            current = shift.Kind == RowShiftKind.Insert
                ? r < shift.FirstRow ? r : r <= shift.LastRow ? null : r - shift.Count
                : r < shift.FirstRow ? r : r + shift.Count;
        }

        return current;
    }

    /// <summary>
    /// The offset by which a range as read moves with its card: when its rows lie inside one kept card's rows, the card's move; otherwise
    /// null, and the range is moved by the shifts.
    /// </summary>
    public int? CardOffset(int firstRow, int lastRow) =>
        slotOf.TryGetValue(firstRow, out var slot) && lastRow <= slot.Last && kept.TryGetValue(firstRow, out var moved)
            ? moved - firstRow
            : null;

    /// <summary>True when the row as read was a removed card's row.</summary>
    public bool IsRemoved(int row) => removed.Contains(row);

    /// <summary>A range (merge, sqref) moved by the shifts; null when every cell of it is deleted.</summary>
    public string? ShiftRange(string range)
    {
        string? current = range;
        foreach (var shift in Shifts)
        {
            if (current is null)
            {
                break;
            }

            current = FormulaReferences.ShiftRange(current, shift);
        }

        return current;
    }

    /// <summary>
    /// A range that belongs to the rows it covers (a merge, a hyperlink, a comment): moved with its card when it lies inside one kept card,
    /// otherwise by the shifts; null when it is deleted.
    /// </summary>
    public string? MoveRange(string range)
    {
        if (CellReference.TryParseRange(range, out var area) && CardOffset(area.FirstRow, area.LastRow) is { } offset)
        {
            return offset == 0 ? range : Name(area with { FirstRow = area.FirstRow + offset, LastRow = area.LastRow + offset });
        }

        return ShiftRange(range);
    }

    /// <summary>
    /// A row point (a drawing marker, a page break) moved by the shifts; a point in deleted rows goes to the first row after them.
    /// </summary>
    public int ShiftPointClamped(int row)
    {
        var current = row;
        foreach (var shift in Shifts)
        {
            if (shift.Kind == RowShiftKind.Insert)
            {
                current = current >= shift.FirstRow ? current + shift.Count : current;
            }
            else
            {
                current = current < shift.FirstRow ? current : current > shift.LastRow ? current - shift.Count : shift.FirstRow;
            }
        }

        return Math.Min(current, CellReference.MaxRow);
    }

    public static string Name(MergeRange range) =>
        range.FirstRow == range.LastRow && range.FirstColumn == range.LastColumn
            ? CellIndex.Name(range.FirstRow, range.FirstColumn)
            : CellIndex.Name(range.FirstRow, range.FirstColumn) + ":" + CellIndex.Name(range.LastRow, range.LastColumn);

    private int? ShiftPoint(int row)
    {
        int? current = row;
        foreach (var shift in Shifts)
        {
            if (current is not { } r)
            {
                break;
            }

            current = shift.Kind == RowShiftKind.Insert
                ? r >= shift.FirstRow ? r + shift.Count : r
                : r < shift.FirstRow ? r : r > shift.LastRow ? r - shift.Count : null;
        }

        return current is > CellReference.MaxRow ? null : current;
    }

    /// <summary>One card's rows after the change.</summary>
    /// <param name="FirstRow">Its first row.</param>
    /// <param name="Rows">How many rows it takes.</param>
    /// <param name="SourceFirstRow">The first row it was read from when it keeps its rows; null for a card in new rows.</param>
    /// <param name="OwnRows">How many of its rows are its rows as read (0 for a card in new rows); the others are new.</param>
    internal sealed record CardRows(int FirstRow, int Rows, int? SourceFirstRow, int OwnRows)
    {
        public bool IsKept => SourceFirstRow is not null;
    }
}
