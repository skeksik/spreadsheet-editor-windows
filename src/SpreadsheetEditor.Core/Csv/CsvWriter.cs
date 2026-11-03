using System.Text;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;

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
    /// <summary>Writes the file on the thread pool (docs/spec.md §9.53).</summary>
    /// <inheritdoc cref="Write(Sheet, string, CsvFormat)"/>
    public static Task WriteAsync(Sheet sheet, string path, CsvFormat format, CancellationToken cancellationToken = default) =>
        Task.Run(() => Write(sheet, path, format), cancellationToken);

    /// <summary>
    /// Writes <paramref name="sheet"/> to the file at <paramref name="path"/>, replacing it: into a temporary file in the target's folder,
    /// which then replaces the target, as <see cref="Xlsx.XlsxWriter"/> does, so a failed write leaves the target as it was.
    /// </summary>
    /// <exception cref="SpreadsheetWriteException">
    /// The file cannot be written (<see cref="SpreadsheetWriteError.AccessDenied"/>, <see cref="SpreadsheetWriteError.WriteFailed"/>), or a
    /// value has a character the format's encoding cannot hold, such as a lone surrogate (<see cref="SpreadsheetWriteError.WriteFailed"/>).
    /// </exception>
    public static void Write(Sheet sheet, string path, CsvFormat format)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(format);
        string? temporary = null;
        try
        {
            var full = Path.GetFullPath(path);
            var folder = Path.GetDirectoryName(full) ?? ".";
            temporary = Path.Combine(folder, $"{Path.GetFileName(full)}.{Guid.NewGuid():N}.tmp");
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 64 * 1024))
            {
                Write(sheet, stream, format);
                stream.Flush(flushToDisk: true);
            }

            File.Move(temporary, full, overwrite: true);
            temporary = null;
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new SpreadsheetWriteException(SpreadsheetWriteError.AccessDenied, $"The file cannot be written: {ex.Message}", ex);
        }
        catch (IOException ex)
        {
            throw new SpreadsheetWriteException(SpreadsheetWriteError.WriteFailed, $"The file could not be written: {ex.Message}", ex);
        }
        catch (EncoderFallbackException ex)
        {
            throw new SpreadsheetWriteException(
                SpreadsheetWriteError.WriteFailed, $"A value has a character the file's encoding cannot hold: {ex.Message}", ex);
        }
        finally
        {
            if (temporary is not null)
            {
                TryDelete(temporary);
            }
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Left behind; the save has already failed with the error that matters.
        }
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
