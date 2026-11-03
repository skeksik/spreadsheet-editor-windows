namespace SpreadsheetEditor.Core.Csv;

/// <summary>
/// How a .csv file is written: its separator, encoding and whether the last line ends with a line break. A sheet read from a .csv keeps
/// the format it was read with, so it is written back the same way; any other sheet is exported with <see cref="Default"/>
/// (docs/spec.md §5.9, §9.15).
/// </summary>
/// <param name="Separator">The separator between values.</param>
/// <param name="Encoding">The text encoding and its byte order mark.</param>
/// <param name="FinalLineBreak">
/// Whether the last line ends with CRLF. A file read without one (as Notepad saves it) is written back without one.
/// </param>
public sealed record CsvFormat(CsvSeparator Separator, CsvEncoding Encoding, bool FinalLineBreak = true)
{
    /// <summary>Comma-separated, UTF-8 without a byte order mark, a line break after the last line.</summary>
    public static CsvFormat Default { get; } = new(CsvSeparator.Comma, CsvEncoding.Utf8);
}
