using System.Text;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Formulas;

/// <summary>
/// An A1 reference without its sheet: a cell, a range of cells, whole rows or whole columns. <see cref="Last"/> equals
/// <see cref="First"/> when <see cref="IsRange"/> is false. The parts keep the order they were written in.
/// </summary>
internal readonly record struct A1Area(A1AreaKind Kind, bool IsRange, A1Part First, A1Part Last)
{
    /// <summary>
    /// Parses an area starting at <paramref name="start"/>: <c>A1</c>, <c>A1:B2</c>, <c>A:B</c> or <c>1:2</c>, with <c>$</c> before
    /// any column or row. The character after the area must not continue a name, open a call or a bracket, or start a sheet name
    /// (<c>!</c>), so that <c>LOG10(</c>, <c>A1B</c> and <c>Q1!</c> are not references.
    /// </summary>
    public static bool TryParse(string text, int start, out A1Area area, out int end)
    {
        area = default;
        end = start;
        if (TryParseCell(text, start, out var first, out var afterFirst))
        {
            if (afterFirst < text.Length && text[afterFirst] == ':'
                && TryParseCell(text, afterFirst + 1, out var last, out var afterLast) && IsBoundary(text, afterLast))
            {
                area = new A1Area(A1AreaKind.Cells, true, first, last);
                end = afterLast;
                return true;
            }
            if (IsBoundary(text, afterFirst))
            {
                area = new A1Area(A1AreaKind.Cells, false, first, first);
                end = afterFirst;
                return true;
            }
            return false;
        }
        var i = start;
        if (TryParseColumn(text, ref i, out var c1, out var c1Abs) && i < text.Length && text[i] == ':')
        {
            i++;
            if (TryParseColumn(text, ref i, out var c2, out var c2Abs) && IsBoundary(text, i))
            {
                area = new A1Area(A1AreaKind.Columns, true, new A1Part(0, false, c1, c1Abs), new A1Part(0, false, c2, c2Abs));
                end = i;
                return true;
            }
            return false;
        }
        i = start;
        if (TryParseRow(text, ref i, out var r1, out var r1Abs) && i < text.Length && text[i] == ':')
        {
            i++;
            if (TryParseRow(text, ref i, out var r2, out var r2Abs) && IsBoundary(text, i))
            {
                area = new A1Area(A1AreaKind.Rows, true, new A1Part(r1, r1Abs, 0, false), new A1Part(r2, r2Abs, 0, false));
                end = i;
                return true;
            }
        }
        return false;
    }

    /// <summary>The area as A1 text, for example <c>$A1:B$2</c>.</summary>
    public override string ToString()
    {
        var builder = new StringBuilder();
        Append(builder, First);
        if (IsRange)
        {
            builder.Append(':');
            Append(builder, Last);
        }
        return builder.ToString();
    }

    /// <summary>Whether the character at <paramref name="index"/> may follow a reference.</summary>
    private static bool IsBoundary(string text, int index)
    {
        if (index >= text.Length)
        {
            return true;
        }
        var c = text[index];
        return !(FormulaTokenizer.IsNameChar(c) || c is '(' or '[' or '!' or '$' or '\'');
    }

    private static void Append(StringBuilder builder, A1Part part)
    {
        if (part.Column > 0)
        {
            if (part.ColumnAbsolute)
            {
                builder.Append('$');
            }
            builder.Append(CellReference.ColumnName(part.Column));
        }
        if (part.Row > 0)
        {
            if (part.RowAbsolute)
            {
                builder.Append('$');
            }
            builder.Append(part.Row);
        }
    }

    private static bool TryParseCell(string text, int start, out A1Part part, out int end)
    {
        part = default;
        end = start;
        var i = start;
        if (!TryParseColumn(text, ref i, out var column, out var columnAbsolute)
            || !TryParseRow(text, ref i, out var row, out var rowAbsolute))
        {
            return false;
        }
        part = new A1Part(row, rowAbsolute, column, columnAbsolute);
        end = i;
        return true;
    }

    /// <summary>
    /// <c>$?</c> and one to three ASCII letters up to <c>XFD</c>; the whole run of letters must fit, so <c>ABCD</c> is not a column.
    /// Leaves <paramref name="index"/> unchanged on failure.
    /// </summary>
    private static bool TryParseColumn(string text, ref int index, out int column, out bool absolute)
    {
        column = 0;
        var i = index;
        absolute = i < text.Length && text[i] == '$';
        if (absolute)
        {
            i++;
        }
        var start = i;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
        {
            if (i - start == 3)
            {
                return false;
            }
            column = (column * 26) + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            i++;
        }
        if (i == start || column > CellReference.MaxColumn)
        {
            return false;
        }
        index = i;
        return true;
    }

    /// <summary>
    /// <c>$?</c> and a row number from 1 to 1,048,576 without leading zeros; the whole run of digits must fit. Leaves
    /// <paramref name="index"/> unchanged on failure.
    /// </summary>
    private static bool TryParseRow(string text, ref int index, out int row, out bool absolute)
    {
        row = 0;
        var i = index;
        absolute = i < text.Length && text[i] == '$';
        if (absolute)
        {
            i++;
        }
        var start = i;
        if (i >= text.Length || text[i] is < '1' or > '9')
        {
            return false;
        }
        while (i < text.Length && char.IsAsciiDigit(text[i]))
        {
            if (i - start == 7)
            {
                return false;
            }
            row = (row * 10) + (text[i] - '0');
            i++;
        }
        if (row > CellReference.MaxRow)
        {
            return false;
        }
        index = i;
        return true;
    }
}
