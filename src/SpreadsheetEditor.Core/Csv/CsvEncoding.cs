namespace SpreadsheetEditor.Core.Csv;

/// <summary>The text encodings a .csv file may use (docs/spec.md §9.15). No legacy code pages.</summary>
public enum CsvEncoding
{
    /// <summary>UTF-8 without a byte order mark.</summary>
    Utf8,

    /// <summary>UTF-8 with a byte order mark (EF BB BF).</summary>
    Utf8WithBom,

    /// <summary>UTF-16 little-endian, with a byte order mark (FF FE).</summary>
    Utf16LittleEndian,

    /// <summary>UTF-16 big-endian, with a byte order mark (FE FF).</summary>
    Utf16BigEndian,
}
