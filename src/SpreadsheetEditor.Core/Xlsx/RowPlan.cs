using SpreadsheetEditor.Core.Formulas;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Where each worksheet row of a sheet read from <c>.xlsx</c> goes when its table is written (docs/spec.md §5.3, §9.18, §9.20, §9.22): the
/// cards in the model's order, each with its own whole rows, new cards in new rows, removed cards' rows deleted, and everything below the
/// table moved so that it keeps its place relative to the table's end.
/// </summary>
/// <remarks>
/// <para>The change is also expressed as a sequence of <see cref="RowShift"/>s: first the removed cards' rows are deleted, bottom up, then
/// the new rows are inserted, top down, at their final rows (the table name row first). Reordering the cards that stay is not in that
/// sequence. What belongs to a row (its cells, height and style, the merges, hyperlinks, comments and pictures inside one card's rows)
/// moves with its card (<see cref="Map"/>, <see cref="MoveRange"/>); a <c>sqref</c> (conditional formatting, data validation, ignored
/// errors, protected ranges) and a table part's range go through the shifts alone (<see cref="ShiftRange"/>), since one range may cover
/// several cards.</para>
/// <para>References in formulas, defined names (the print area included), hyperlink locations and charts go through
/// <see cref="MoveArea"/> (docs/spec.md §9.20; owner, 2026-11-03, and manager within the owner's answer):</para>
/// <list type="number">
/// <item>A cell follows its card, as a reference to a cut cell follows it in Excel: a reference to Cherry's quantity points at Cherry's
/// row after the save; one to a removed card's cell becomes <c>#REF!</c>; one outside the table moves with the shifts.</item>
/// <item>A range whose rows lie inside one kept card's rows moves with the card, as its merges do.</item>
/// <item>Any other range (endpoints in different cards, or one in the table and one outside it) keeps covering the same block of rows:
/// it goes through the shifts alone, growing by rows inserted inside it and shrinking by rows deleted from it; a range whose rows are all
/// deleted becomes <c>#REF!</c>. Not verified against Excel: whether Excel's Insert Cut Cells moves a range's endpoint that is itself a
/// moved cell (Microsoft's "Move or copy cells, rows, and columns" and "Move or copy a formula", read 2026-11-03, do not say). This rule
/// keeps a total over the table a total over the table instead.</item>
/// <item>As an Excel Table's references do: a range whose bottom row as read is the table's last card row as read ends at the table's
/// last card row after the save, so cards added at the end count; a range from the first card row to the last (or below) as read starts
/// at the first card row after the save, so cards added at the top count too (manager, 2026-11-03, within the owner's answer). When the
/// table has no card rows left, rule 3 alone applies.</item>
/// <item>Whole columns never change; whole rows (<c>2:6</c>) go through the shifts alone.</item>
/// </list>
/// <para>A card read from the file keeps its rows and gets more when its <see cref="Card.SourceRowSpan"/> asks for more; any other card
/// (new or duplicated, §9.22) gets <see cref="Card.SourceRowSpan"/> new rows.</para>
/// </remarks>
internal sealed class RowPlan
{
    private readonly Dictionary<int, int> kept = [];
    private readonly Dictionary<int, int> keptBack = [];
    private readonly Dictionary<int, (int First, int Last)> slotOf = [];
    private readonly HashSet<int> removed = [];
    private int dataLastRow;
    private bool insertsName;
    private (int First, int Last)? readData;

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

        // The final layout: each card's rows in order. A kept card's own rows keep their content; its extra rows and new cards' rows are
        // inserted.
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
                }
            }

            for (var k = own; k < rows; k++)
            {
                inserted.Add(cursor + k);
            }

            plan.Cards.Add(new CardRows(cursor, rows, slot is { } i ? slots[i].FirstRow : null, own));
            cursor += rows;
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
        plan.readData = slots.Count > 0 ? (slots[0].FirstRow, slots[^1].FirstRow + slots[^1].RowSpan - 1) : null;
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

    /// <summary>
    /// Where a reference to this sheet's rows as read points after the change (docs/spec.md §9.20); null when its cells are deleted
    /// (<c>#REF!</c>). The rules are in the class's remarks.
    /// </summary>
    internal A1Area? MoveArea(A1Area area)
    {
        if (area.Kind == A1AreaKind.Columns)
        {
            return area;
        }

        if (area.Kind == A1AreaKind.Cells && !area.IsRange)
        {
            // A cell follows its card, or goes with its deleted row.
            return Map(area.First.Row) is { } row ? WithRows(area, row, row) : null;
        }

        var top = Math.Min(area.First.Row, area.Last.Row);
        var bottom = Math.Max(area.First.Row, area.Last.Row);
        if (area.Kind == A1AreaKind.Cells && CardOffset(top, bottom) is { } offset)
        {
            return offset == 0 ? area : WithRows(area, top + offset, bottom + offset);
        }

        A1Area? current = area;
        foreach (var shift in Shifts)
        {
            if (current is not { } a)
            {
                return null;
            }

            current = FormulaReferences.ShiftArea(a, shift);
        }

        if (current is not { } shifted || area.Kind != A1AreaKind.Cells || readData is not { } read || dataLastRow < FirstDataRow)
        {
            return current;
        }

        // A range that reached the table's last row reaches it after the change (cards added at the end count); one that covered every
        // card covers every card (cards added at the top count too).
        var newTop = Math.Min(shifted.First.Row, shifted.Last.Row);
        var newBottom = Math.Max(shifted.First.Row, shifted.Last.Row);
        if (bottom == read.Last)
        {
            newBottom = Math.Max(newBottom, dataLastRow);
        }

        if (top == read.First && bottom >= read.Last)
        {
            newTop = Math.Min(newTop, FirstDataRow);
        }

        return WithRows(shifted, newTop, newBottom);
    }

    /// <summary>The area with these top and bottom rows, keeping the order its parts were written in.</summary>
    private static A1Area WithRows(A1Area area, int top, int bottom)
    {
        if (!area.IsRange)
        {
            return area with { First = area.First with { Row = top }, Last = area.Last with { Row = top } };
        }

        var firstIsTop = area.First.Row <= area.Last.Row;
        return area with
        {
            First = area.First with { Row = firstIsTop ? top : bottom },
            Last = area.Last with { Row = firstIsTop ? bottom : top },
        };
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
