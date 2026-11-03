using System.Text;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Csv;

/// <summary>
/// Reads a .csv file into a workbook of one sheet (docs/spec.md §4.7, §9.4, §9.14, §9.15).
/// <list type="bullet">
/// <item>Encoding: from the byte order mark (UTF-8, UTF-16 little- or big-endian); without one, UTF-8. Bytes that are not valid in that
/// encoding, a null character, or a UTF-32 byte order mark fail with <see cref="CsvReadFailure.UnsupportedEncoding"/>.</item>
/// <item>Separator: detected from the header line (comma, semicolon or tab; see <see cref="CsvParser.DetectSeparator"/>).</item>
/// <item>Values: RFC 4180 quoting, values may span lines; blank lines are skipped; every value is text, kept exactly as read.</item>
/// <item>The first record is the header; column names are kept as read, blank and repeated ones included (§9.36–9.38).</item>
/// <item>One card per record after the header; a header-only file gives one empty card. A short record gets empty values; a long one
/// adds columns with empty names, so nothing is dropped.</item>
/// </list>
/// </summary>
public static class CsvReader
{
    /// <summary>Reads the file at <paramref name="path"/>; the sheet is named after the file name with its extension.</summary>
    public static CsvReadResult Read(string path)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 64 * 1024, FileOptions.SequentialScan);
        var result = Read(stream, Path.GetFileName(path));
        result.Workbook.SourcePath = path;
        return result;
    }

    /// <summary>Reads .csv bytes from <paramref name="stream"/> into a sheet named <paramref name="sheetName"/>.</summary>
    public static CsvReadResult Read(Stream stream, string sheetName)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(sheetName);
        if (stream.CanSeek)
        {
            return ReadSeekable(stream, sheetName);
        }
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        copy.Position = 0;
        return ReadSeekable(copy, sheetName);
    }

    private static CsvReadResult ReadSeekable(Stream stream, string sheetName)
    {
        var encoding = DetectEncoding(stream);
        using var text = new StreamReader(stream, CsvText.GetEncoding(encoding), detectEncodingFromByteOrderMarks: false, 64 * 1024, leaveOpen: true);
        var parser = new CsvParser(text);
        try
        {
            var separator = parser.DetectSeparator();
            var header = parser.ReadRecord() ?? throw new CsvReadException(CsvReadFailure.Empty, 0, "The file is empty: it has no header line.");
            var records = new List<List<string>>();
            var columnCount = header.Count;
            while (parser.ReadRecord() is { } record)
            {
                records.Add(record);
                columnCount = Math.Max(columnCount, record.Count);
            }
            var sheet = new Sheet(sheetName);
            for (var i = 0; i < columnCount; i++)
            {
                sheet.Columns.Add(new Column(i < header.Count ? header[i] : ""));
            }
            foreach (var record in records)
            {
                sheet.Cards.Add(MakeCard(record, columnCount));
            }
            if (sheet.Cards.Count == 0)
            {
                sheet.Cards.Add(MakeCard([], columnCount));
            }
            var workbook = new Workbook();
            workbook.Sheets.Add(sheet);
            return new CsvReadResult(workbook, new CsvFormat(CsvText.FromChar(separator), encoding));
        }
        catch (DecoderFallbackException e)
        {
            throw new CsvReadException(
                CsvReadFailure.UnsupportedEncoding,
                parser.LineNumber,
                "The file is not UTF-8, nor UTF-16 with a byte order mark.",
                e);
        }
    }

    /// <summary>Reads the byte order mark, leaves the stream just after it and returns the encoding it names (UTF-8 without one).</summary>
    private static CsvEncoding DetectEncoding(Stream stream)
    {
        var start = stream.Position;
        Span<byte> head = stackalloc byte[4];
        var count = stream.ReadAtLeast(head, head.Length, throwOnEndOfStream: false);
        head = head[..count];
        if (head is [0xFF, 0xFE, 0x00, 0x00] or [0x00, 0x00, 0xFE, 0xFF])
        {
            throw new CsvReadException(CsvReadFailure.UnsupportedEncoding, 0, "The file is UTF-32, which is not supported.");
        }
        var (encoding, length) = head switch
        {
            [0xEF, 0xBB, 0xBF, ..] => (CsvEncoding.Utf8WithBom, 3),
            [0xFF, 0xFE, ..] => (CsvEncoding.Utf16LittleEndian, 2),
            [0xFE, 0xFF, ..] => (CsvEncoding.Utf16BigEndian, 2),
            _ => (CsvEncoding.Utf8, 0),
        };
        stream.Position = start + length;
        return encoding;
    }

    private static Card MakeCard(List<string> values, int columnCount)
    {
        var card = new Card();
        for (var i = 0; i < columnCount; i++)
        {
            var text = i < values.Count ? values[i] : "";
            card.Entries.Add(new Entry
            {
                Text = text,
                Original = text.Length == 0 ? null : new CellValue(CellValueKind.Text, text),
            });
        }
        return card;
    }
}
