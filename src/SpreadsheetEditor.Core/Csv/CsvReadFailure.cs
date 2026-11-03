namespace SpreadsheetEditor.Core.Csv;

/// <summary>Why a .csv file could not be read.</summary>
public enum CsvReadFailure
{
    /// <summary>The file has no header line: it is empty, holds only a byte order mark, or only blank lines.</summary>
    Empty,

    /// <summary>The file is not UTF-8, nor UTF-16 with a byte order mark (a legacy code page, UTF-16 without a byte order mark, or UTF-32).</summary>
    UnsupportedEncoding,

    /// <summary>A quoted value is not closed before the end of the file.</summary>
    UnterminatedQuote,
}
