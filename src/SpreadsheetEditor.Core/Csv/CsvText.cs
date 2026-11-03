using System.Text;

namespace SpreadsheetEditor.Core.Csv;

/// <summary>The characters and encodings behind <see cref="CsvSeparator"/> and <see cref="CsvEncoding"/>.</summary>
internal static class CsvText
{
    public static char ToChar(CsvSeparator separator) => separator switch
    {
        CsvSeparator.Comma => ',',
        CsvSeparator.Semicolon => ';',
        CsvSeparator.Tab => '\t',
        _ => throw new ArgumentOutOfRangeException(nameof(separator)),
    };

    public static CsvSeparator FromChar(char separator) => separator switch
    {
        ',' => CsvSeparator.Comma,
        ';' => CsvSeparator.Semicolon,
        '\t' => CsvSeparator.Tab,
        _ => throw new ArgumentOutOfRangeException(nameof(separator)),
    };

    /// <summary>The encoding without its byte order mark, throwing on bytes or characters it cannot convert.</summary>
    public static Encoding GetEncoding(CsvEncoding encoding) => encoding switch
    {
        CsvEncoding.Utf8 or CsvEncoding.Utf8WithBom => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
        CsvEncoding.Utf16LittleEndian => new UnicodeEncoding(bigEndian: false, byteOrderMark: false, throwOnInvalidBytes: true),
        CsvEncoding.Utf16BigEndian => new UnicodeEncoding(bigEndian: true, byteOrderMark: false, throwOnInvalidBytes: true),
        _ => throw new ArgumentOutOfRangeException(nameof(encoding)),
    };

    private static readonly byte[] Utf8ByteOrderMark = [0xEF, 0xBB, 0xBF];
    private static readonly byte[] Utf16LittleEndianByteOrderMark = [0xFF, 0xFE];
    private static readonly byte[] Utf16BigEndianByteOrderMark = [0xFE, 0xFF];

    /// <summary>The byte order mark written at the start of the file, empty for UTF-8 without one.</summary>
    public static ReadOnlySpan<byte> GetByteOrderMark(CsvEncoding encoding) => encoding switch
    {
        CsvEncoding.Utf8 => ReadOnlySpan<byte>.Empty,
        CsvEncoding.Utf8WithBom => Utf8ByteOrderMark,
        CsvEncoding.Utf16LittleEndian => Utf16LittleEndianByteOrderMark,
        CsvEncoding.Utf16BigEndian => Utf16BigEndianByteOrderMark,
        _ => throw new ArgumentOutOfRangeException(nameof(encoding)),
    };
}
