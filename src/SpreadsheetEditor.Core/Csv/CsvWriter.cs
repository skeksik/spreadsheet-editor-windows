using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Csv;

/// <summary>
/// Writes one sheet as a .csv file (docs/spec.md §5.9, §9.14, §9.34): the caller passes the sheet (the app passes the selected one) and its
/// format, the format the sheet was read with or <see cref="CsvFormat.Default"/>.
/// <list type="bullet">
/// <item>A header line of the column names, unless the sheet's headers are hidden (<see cref="Sheet.HeaderRow"/> is null).</item>
/// <item>One line per card with each entry's text; pictures, styles, row spans and the table name are not written. A sheet whose only
/// card is empty writes no card line, so a header-only file is saved as it was read.</item>
/// <item>A value is quoted only when it holds the separator, a quote or a line break, and a line holding one empty value is written as
/// <c>""</c>, so that it is not read back as a blank line.</item>
/// <item>CRLF after every line, the last one included; the format's encoding, with its byte order mark if it has one.</item>
/// </list>
/// </summary>
public static class CsvWriter
{
    /// <summary>Writes <paramref name="sheet"/> to the file at <paramref name="path"/>, replacing it.</summary>
    public static void Write(Sheet sheet, string path, CsvFormat format)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 64 * 1024);
        Write(sheet, stream, format);
    }

    /// <summary>Writes <paramref name="sheet"/> to <paramref name="stream"/>, which stays open.</summary>
    public static void Write(Sheet sheet, Stream stream, CsvFormat format)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(format);
        var separator = CsvText.ToChar(format.Separator);
        stream.Write(CsvText.GetByteOrderMark(format.Encoding));
        using var writer = new StreamWriter(stream, CsvText.GetEncoding(format.Encoding), 64 * 1024, leaveOpen: true);
        var columnCount = sheet.Columns.Count;
        if (columnCount == 0)
        {
            return;
        }
        char[] needsQuotes = [separator, '"', '\r', '\n'];
        var values = new string[columnCount];
        if (sheet.HeaderRow is not null)
        {
            for (var i = 0; i < columnCount; i++)
            {
                values[i] = sheet.Columns[i].Name;
            }
            WriteLine(writer, values, separator, needsQuotes);
        }
        if (sheet.Cards.Count == 1 && IsEmpty(sheet.Cards[0]))
        {
            return;
        }
        foreach (var card in sheet.Cards)
        {
            for (var i = 0; i < columnCount; i++)
            {
                values[i] = i < card.Entries.Count ? card.Entries[i].Text : "";
            }
            WriteLine(writer, values, separator, needsQuotes);
        }
    }

    private static bool IsEmpty(Card card)
    {
        foreach (var entry in card.Entries)
        {
            if (entry.Text.Length > 0)
            {
                return false;
            }
        }
        return true;
    }

    private static void WriteLine(StreamWriter writer, string[] values, char separator, char[] needsQuotes)
    {
        if (values.Length == 1 && values[0].Length == 0)
        {
            writer.Write("\"\"");
        }
        else
        {
            for (var i = 0; i < values.Length; i++)
            {
                if (i > 0)
                {
                    writer.Write(separator);
                }
                var value = values[i];
                if (value.IndexOfAny(needsQuotes) >= 0)
                {
                    writer.Write('"');
                    writer.Write(value.Replace("\"", "\"\"", StringComparison.Ordinal));
                    writer.Write('"');
                }
                else
                {
                    writer.Write(value);
                }
            }
        }
        writer.Write("\r\n");
    }
}
