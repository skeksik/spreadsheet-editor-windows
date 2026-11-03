using System.Text;
using System.Text.RegularExpressions;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Formulas;

/// <summary>
/// Rewrites the A1 references in formula text when worksheet rows are inserted or deleted, and translates a shared formula to another
/// cell, and renames a sheet in its references. Pure functions over text: no workbook, no evaluation. Text in double quotes, function names, defined names, error literals and
/// bracketed text are copied unchanged, and so is everything else that is not a reference.
/// </summary>
/// <remarks>
/// <para>Left unchanged on purpose:</para>
/// <list type="bullet">
/// <item>Structured (table) references such as <c>Table1[Price]</c> or <c>[@Price]</c>: they name a table's columns and regions, not
/// rows, so inserting rows changes the table's own range (see <see cref="ShiftRange"/>), not the formula text. This is reasoning, not
/// verified against a source.</item>
/// <item>References to several sheets (<c>Sheet1:Sheet3!A1</c>), whose span depends on the workbook's sheet order, and references to
/// other workbooks (<c>[1]Sheet1!A1</c>), when shifting rows. <see cref="Translate"/> moves their relative parts like any other.</item>
/// <item>R1C1 notation, which is not parsed.</item>
/// </list>
/// <para>Not verified against Excel: the text written for a reference to a deleted cell is <c>#REF!</c> after the sheet prefix when it
/// has one (<c>Sheet1!#REF!</c>), as Excel shows it; a range that ends on the sheet's last row keeps ending there; a range whose
/// end is pushed past the last row by an insertion is cut at the last row; a reference that <see cref="Translate"/> moves off the sheet
/// becomes <c>#REF!</c>.</para>
/// </remarks>
public static partial class FormulaReferences
{
    private const string RefError = "#REF!";

    /// <summary>
    /// Returns <paramref name="formula"/> with every sheet prefix naming <paramref name="oldName"/> (compared with
    /// <see cref="SheetNames.Comparer"/>) changed to <paramref name="newName"/>: the prefixes of references (<c>Sheet1!A1</c>), of
    /// sheet-level names and errors (<c>Sheet1!MyName</c>, <c>Sheet1!#REF!</c>), and either end of a reference to several sheets
    /// (<c>Sheet1:Sheet3!A1</c>). A changed prefix is written with quotes when the name needs them (<see cref="NeedsQuotes"/>), with
    /// any apostrophe doubled; a reference to several sheets is quoted as a whole. Unqualified references, references to other
    /// workbooks (<c>[1]Sheet1!A1</c>), text in double quotes and bracketed text are copied unchanged, and so are prefixes that do not
    /// name the old sheet.
    /// </summary>
    /// <param name="formula">The formula text, with or without its leading <c>=</c>.</param>
    /// <param name="oldName">The sheet's name before the rename.</param>
    /// <param name="newName">The sheet's new name.</param>
    public static string RenameSheet(string formula, string oldName, string newName)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ArgumentException.ThrowIfNullOrEmpty(oldName);
        ArgumentException.ThrowIfNullOrEmpty(newName);
        if (string.Equals(oldName, newName, StringComparison.Ordinal))
        {
            return formula;
        }
        StringBuilder? builder = null;
        var copied = 0;
        foreach (var token in FormulaTokenizer.Tokenize(formula))
        {
            if (token.SheetName is not { } name || token.PrefixLength == 0 || token.IsExternal)
            {
                continue;
            }
            string? prefix = null;
            var colon = name.IndexOf(':', StringComparison.Ordinal);
            if (token.IsMultiSheetOrExternal && colon >= 0)
            {
                var first = name[..colon];
                var last = name[(colon + 1)..];
                var renamedFirst = SheetNames.Comparer.Equals(first, oldName);
                var renamedLast = SheetNames.Comparer.Equals(last, oldName);
                if (renamedFirst || renamedLast)
                {
                    first = renamedFirst ? newName : first;
                    last = renamedLast ? newName : last;
                    var span = first + ":" + last;
                    prefix = NeedsQuotes(first) || NeedsQuotes(last) ? Quote(span) : span;
                }
            }
            else if (SheetNames.Comparer.Equals(name, oldName))
            {
                prefix = NeedsQuotes(newName) ? Quote(newName) : newName;
            }
            if (prefix is null)
            {
                continue;
            }
            builder ??= new StringBuilder(formula.Length + newName.Length);
            builder.Append(formula, copied, token.Start - copied).Append(prefix).Append('!');
            copied = token.Start + token.PrefixLength;
        }
        return builder is null ? formula : builder.Append(formula, copied, formula.Length - copied).ToString();
    }

    /// <summary>
    /// Whether a sheet name must be quoted (<c>'My sheet'!A1</c>) in a reference. Not verified against ECMA-376's formula grammar
    /// (Part 1 §18.17) or Excel; the rule errs toward quoting, which is always allowed: a name is left bare only when it is made of
    /// letters, digits, <c>_</c> and <c>.</c>, does not start with a digit or <c>.</c>, and does not read as an A1 cell reference
    /// (1–3 letters then digits, as <c>A1</c> or <c>Abc12</c>), an R1C1 reference (<c>R</c>, <c>C</c>, <c>R2</c>, <c>C3</c>,
    /// <c>RC</c>, <c>R1C1</c>, in either case) or a boolean (<c>TRUE</c>, <c>FALSE</c>).
    /// </summary>
    internal static bool NeedsQuotes(string name)
    {
        if (name.Length == 0 || char.IsAsciiDigit(name[0]) || name[0] == '.')
        {
            return true;
        }
        foreach (var c in name)
        {
            if (!(char.IsLetterOrDigit(c) || c is '_' or '.'))
            {
                return true;
            }
        }
        return CellLike().IsMatch(name)
            || name.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
            || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase);
    }

    private static string Quote(string name) => "'" + name.Replace("'", "''", StringComparison.Ordinal) + "'";

    /// <summary>An A1 cell reference (1–3 letters then digits) or an R1C1 reference, in either case.</summary>
    [GeneratedRegex(@"^(?:[A-Za-z]{1,3}[0-9]+|[Rr][0-9]*(?:[Cc][0-9]*)?|[Cc][0-9]*)$", RegexOptions.CultureInvariant)]
    private static partial Regex CellLike();

    /// <summary>
    /// Returns <paramref name="formula"/> with its references to <see cref="RowShift.Sheet"/> moved for the inserted or deleted rows, as
    /// a spreadsheet program does: relative and absolute rows both move, a range spanning an insertion grows, a range losing some
    /// rows shrinks, and a reference whose cells are all deleted becomes <c>#REF!</c>. Whole columns (<c>A:C</c>) never change.
    /// </summary>
    /// <param name="formula">The formula text, with or without its leading <c>=</c>.</param>
    /// <param name="formulaSheet">The sheet the formula lives on, which unqualified references point at; <see langword="null"/> when
    /// it has none, as for a workbook-level defined name, so that only references qualified with the shifted sheet's name move.</param>
    /// <param name="shift">The rows inserted or deleted.</param>
    public static string ShiftRows(string formula, string? formulaSheet, RowShift shift)
    {
        ArgumentNullException.ThrowIfNull(formula);
        CheckShift(shift);
        var unqualifiedMoves = formulaSheet is not null && SheetNames.Comparer.Equals(formulaSheet, shift.Sheet);
        return Rewrite(formula, token =>
        {
            if (token.IsMultiSheetOrExternal)
            {
                return null;
            }
            var moves = token.SheetName is null ? unqualifiedMoves : SheetNames.Comparer.Equals(token.SheetName, shift.Sheet);
            if (!moves)
            {
                return null;
            }
            var shifted = ShiftArea(token.Area, shift);
            return shifted == token.Area ? null : shifted?.ToString() ?? RefError;
        });
    }

    /// <summary>
    /// Returns a range on <see cref="RowShift.Sheet"/> moved for the inserted or deleted rows, by the rules of
    /// <see cref="ShiftRows"/>: a merge (<c>A1:C3</c>), a table's range, or a list of ranges separated by spaces (a <c>sqref</c>).
    /// Ranges whose cells are all deleted are dropped; when none is left the result is <see langword="null"/>. A print area or a
    /// defined name is a formula: use <see cref="ShiftRows"/> for those.
    /// </summary>
    /// <exception cref="FormatException">A part of <paramref name="range"/> is not an A1 range without a sheet name.</exception>
    public static string? ShiftRange(string range, RowShift shift)
    {
        ArgumentNullException.ThrowIfNull(range);
        CheckShift(shift);
        var parts = new List<string>();
        foreach (var part in range.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (!A1Area.TryParse(part, 0, out var area, out var end) || end != part.Length)
            {
                throw new FormatException($"'{part}' is not a cell range.");
            }
            var shifted = ShiftArea(area, shift);
            if (shifted is { } kept)
            {
                parts.Add(kept == area ? part : kept.ToString());
            }
        }
        return parts.Count == 0 ? null : string.Join(' ', parts);
    }

    /// <summary>
    /// Returns a shared formula as it reads in a cell <paramref name="rowOffset"/> rows and <paramref name="columnOffset"/> columns
    /// away from the cell it was written for: relative rows and columns move by the offsets, absolute (<c>$</c>) ones stay. A
    /// reference moved off the sheet becomes <c>#REF!</c> (not verified).
    /// </summary>
    /// <remarks>
    /// ISO/IEC 29500-1 §18.3.1.40 f (Formula), as quoted in Microsoft Learn's <c>DocumentFormat.OpenXml.Spreadsheet.CellFormula</c>
    /// reference (read 2026-11-03): a shared formula's <c>si</c> "is used to figure what the formula expression should be based on the
    /// cell's relative location to the master formula cell", and "Two formulas are considered to be the same when their respective
    /// representations in R1C1-reference notation, are the same". Keeping the R1C1 form means keeping each relative part's offset
    /// from the cell and each absolute part's position.
    /// </remarks>
    public static string Translate(string formula, int rowOffset, int columnOffset)
    {
        ArgumentNullException.ThrowIfNull(formula);
        if (rowOffset == 0 && columnOffset == 0)
        {
            return formula;
        }
        return Rewrite(formula, token =>
        {
            var area = token.Area;
            if (!TryTranslate(area.First, rowOffset, columnOffset, out var first)
                || !TryTranslate(area.Last, rowOffset, columnOffset, out var last))
            {
                return RefError;
            }
            var moved = area with { First = first, Last = last };
            return moved == area ? null : moved.ToString();
        });
    }

    private static void CheckShift(RowShift shift)
    {
        if (shift.Sheet is null)
        {
            throw new ArgumentException("The row shift has no sheet.", nameof(shift));
        }
    }

    /// <summary>
    /// Copies <paramref name="formula"/>, replacing the area of each reference with <paramref name="replace"/>'s result (the sheet
    /// prefix is kept); <see langword="null"/> keeps the reference's text as it was.
    /// </summary>
    private static string Rewrite(string formula, Func<FormulaToken, string?> replace)
    {
        StringBuilder? builder = null;
        var copied = 0;
        foreach (var token in FormulaTokenizer.Tokenize(formula))
        {
            if (token.Kind != FormulaTokenKind.Reference || replace(token) is not { } area)
            {
                continue;
            }
            builder ??= new StringBuilder(formula.Length + 8);
            var areaStart = token.Start + token.PrefixLength;
            builder.Append(formula, copied, areaStart - copied).Append(area);
            copied = token.End;
        }
        return builder is null ? formula : builder.Append(formula, copied, formula.Length - copied).ToString();
    }

    /// <summary>The area after the shift, or <see langword="null"/> when every cell it covers is deleted.</summary>
    private static A1Area? ShiftArea(A1Area area, RowShift shift)
    {
        if (area.Kind == A1AreaKind.Columns)
        {
            return area;
        }
        var firstIsTop = area.First.Row <= area.Last.Row;
        var top = Math.Min(area.First.Row, area.Last.Row);
        var bottom = Math.Max(area.First.Row, area.Last.Row);
        int newTop, newBottom;
        if (shift.Kind == RowShiftKind.Insert)
        {
            newTop = top >= shift.FirstRow ? top + shift.Count : top;
            newBottom = bottom >= shift.FirstRow ? bottom + shift.Count : bottom;
            if (newTop > RowShift.MaxRow)
            {
                return null;
            }
            if (area.IsRange && (bottom == RowShift.MaxRow || newBottom > RowShift.MaxRow))
            {
                newBottom = RowShift.MaxRow;
            }
            else if (newBottom > RowShift.MaxRow)
            {
                return null;
            }
        }
        else
        {
            if (top >= shift.FirstRow && bottom <= shift.LastRow)
            {
                return null;
            }
            newTop = top < shift.FirstRow ? top : top > shift.LastRow ? top - shift.Count : shift.FirstRow;
            newBottom = bottom < shift.FirstRow ? bottom : bottom > shift.LastRow ? bottom - shift.Count : shift.FirstRow - 1;
            if (area.IsRange && bottom == RowShift.MaxRow)
            {
                newBottom = RowShift.MaxRow;
            }
        }
        if (!area.IsRange)
        {
            return area with { First = area.First with { Row = newTop }, Last = area.Last with { Row = newTop } };
        }
        var first = area.First with { Row = firstIsTop ? newTop : newBottom };
        var last = area.Last with { Row = firstIsTop ? newBottom : newTop };
        return area with { First = first, Last = last };
    }

    private static bool TryTranslate(A1Part part, int rowOffset, int columnOffset, out A1Part moved)
    {
        var row = part.Row == 0 || part.RowAbsolute ? part.Row : part.Row + rowOffset;
        var column = part.Column == 0 || part.ColumnAbsolute ? part.Column : part.Column + columnOffset;
        moved = part with { Row = row, Column = column };
        return (part.Row == 0 || row is >= 1 and <= CellReference.MaxRow)
            && (part.Column == 0 || column is >= 1 and <= CellReference.MaxColumn);
    }
}
