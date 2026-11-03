namespace SpreadsheetEditor.Core.Csv;

/// <summary>
/// How a .csv file is written: its separator and encoding. A sheet read from a .csv keeps the format it was read with; any other sheet is
/// exported with <see cref="Default"/> (docs/spec.md §5.9, §9.15).
/// </summary>
public sealed record CsvFormat(CsvSeparator Separator, CsvEncoding Encoding)
{
    /// <summary>Comma-separated, UTF-8 without a byte order mark.</summary>
    public static CsvFormat Default { get; } = new(CsvSeparator.Comma, CsvEncoding.Utf8);
}
