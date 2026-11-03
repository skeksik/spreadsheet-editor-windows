using System.Text;
using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Csv;

/// <summary>
/// Builds .csv bytes and sheets in code. A sheet is described as text: lines separated by <c>¶</c>, values by <c>|</c>; the first line
/// holds the column names.
/// </summary>
internal static class CsvTestFiles
{
    public const char LineMark = '¶';
    public const char ValueMark = '|';

    public static byte[] Encode(string text, CsvEncoding encoding)
    {
        var (bom, body) = encoding switch
        {
            CsvEncoding.Utf8 => (Array.Empty<byte>(), new UTF8Encoding(false).GetBytes(text)),
            CsvEncoding.Utf8WithBom => (new byte[] { 0xEF, 0xBB, 0xBF }, new UTF8Encoding(false).GetBytes(text)),
            CsvEncoding.Utf16LittleEndian => (new byte[] { 0xFF, 0xFE }, new UnicodeEncoding(false, false).GetBytes(text)),
            CsvEncoding.Utf16BigEndian => (new byte[] { 0xFE, 0xFF }, new UnicodeEncoding(true, false).GetBytes(text)),
            _ => throw new ArgumentOutOfRangeException(nameof(encoding)),
        };
        return [.. bom, .. body];
    }

    public static CsvReadResult Read(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return CsvReader.Read(stream, "data.csv");
    }

    public static CsvReadResult Read(string text) => Read(Encode(text, CsvEncoding.Utf8));

    public static Sheet ReadSheet(string text) => Read(text).Workbook.Sheets[0];

    public static byte[] Write(Sheet sheet, CsvFormat format)
    {
        using var stream = new MemoryStream();
        CsvWriter.Write(sheet, stream, format);
        return stream.ToArray();
    }

    /// <summary>Writes with UTF-8 without a byte order mark and returns the text.</summary>
    public static string WriteText(Sheet sheet, CsvSeparator separator = CsvSeparator.Comma) =>
        Encoding.UTF8.GetString(Write(sheet, new CsvFormat(separator, CsvEncoding.Utf8)));

    public static Sheet MakeSheet(string description)
    {
        var lines = description.Split(LineMark);
        var sheet = new Sheet("Sheet1");
        foreach (var name in lines[0].Split(ValueMark))
        {
            sheet.Columns.Add(new Column(name));
        }
        foreach (var line in lines.Skip(1))
        {
            var card = new Card();
            foreach (var value in line.Split(ValueMark))
            {
                card.Entries.Add(new Entry { Text = value });
            }
            sheet.Cards.Add(card);
        }
        return sheet;
    }

    /// <summary>The sheet's column names and card texts, in <see cref="MakeSheet"/>'s form.</summary>
    public static string Describe(Sheet sheet)
    {
        var lines = new List<string> { string.Join(ValueMark, sheet.Columns.Select(c => c.Name)) };
        lines.AddRange(sheet.Cards.Select(card => string.Join(ValueMark, card.Entries.Select(e => e.Text))));
        return string.Join(LineMark, lines);
    }
}
