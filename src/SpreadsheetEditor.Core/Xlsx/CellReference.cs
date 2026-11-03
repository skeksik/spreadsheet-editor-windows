namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>A1-style cell references: column letters and a row number, both 1-based.</summary>
internal static class CellReference
{
    /// <summary>The last worksheet column, <c>XFD</c>.</summary>
    public const int MaxColumn = 16384;

    /// <summary>The last worksheet row.</summary>
    public const int MaxRow = 1048576;

    /// <summary>Parses a reference such as <c>B12</c> or <c>$B$12</c>.</summary>
    public static bool TryParse(ReadOnlySpan<char> text, out int column, out int row)
    {
        column = 0;
        row = 0;
        var i = 0;
        if (i < text.Length && text[i] == '$')
        {
            i++;
        }
        var letters = 0;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
        {
            column = (column * 26) + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            letters++;
            i++;
            if (letters > 3)
            {
                return false;
            }
        }
        if (letters == 0 || column > MaxColumn)
        {
            return false;
        }
        if (i < text.Length && text[i] == '$')
        {
            i++;
        }
        var digits = 0;
        while (i < text.Length && char.IsAsciiDigit(text[i]))
        {
            row = (row * 10) + (text[i] - '0');
            digits++;
            i++;
            if (digits > 7)
            {
                return false;
            }
        }
        return digits > 0 && i == text.Length && row >= 1 && row <= MaxRow;
    }

    /// <summary>Parses a range such as <c>A2:C3</c>; a single cell is a range of one.</summary>
    public static bool TryParseRange(ReadOnlySpan<char> text, out MergeRange range)
    {
        range = default;
        var colon = text.IndexOf(':');
        if (colon < 0)
        {
            if (!TryParse(text, out var column, out var row))
            {
                return false;
            }
            range = new MergeRange(row, column, row, column);
            return true;
        }
        if (!TryParse(text[..colon], out var c1, out var r1) || !TryParse(text[(colon + 1)..], out var c2, out var r2))
        {
            return false;
        }
        range = new MergeRange(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
        return true;
    }

    /// <summary>The column's letters, for example 1 → <c>A</c>, 28 → <c>AB</c>.</summary>
    public static string ColumnName(int column)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(column, MaxColumn);
        Span<char> buffer = stackalloc char[3];
        var start = buffer.Length;
        while (column > 0)
        {
            var digit = (column - 1) % 26;
            buffer[--start] = (char)('A' + digit);
            column = (column - 1) / 26;
        }
        return new string(buffer[start..]);
    }
}
