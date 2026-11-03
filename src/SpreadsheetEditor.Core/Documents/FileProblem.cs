namespace SpreadsheetEditor.Core.Documents;

/// <summary>
/// What went wrong opening or saving a file, in the terms the app explains to the user (docs/spec.md §10.17); the exception's message is
/// the technical detail shown below it.
/// </summary>
public enum FileProblem
{
    /// <summary>An error the core does not describe; the app shows a general message.</summary>
    Unknown,

    /// <summary>The file or its folder does not exist.</summary>
    NotFound,

    /// <summary>Windows does not allow the file to be opened or written (permissions, a read-only file).</summary>
    AccessDenied,

    /// <summary>Another program has the file open and does not share it (a sharing or lock violation).</summary>
    InUse,

    /// <summary>The file could not be read, for example because another program has it locked.</summary>
    ReadFailed,

    /// <summary>The file could not be written, for example because another program has it locked or the folder does not exist.</summary>
    WriteFailed,

    /// <summary>The file is not an <c>.xlsx</c> workbook.</summary>
    NotAWorkbook,

    /// <summary>The file is an older <c>.xls</c> workbook or a password-protected one.</summary>
    LegacyOrEncrypted,

    /// <summary>The workbook is damaged: a part is missing or its XML cannot be read.</summary>
    Damaged,

    /// <summary>The <c>.csv</c> file is empty.</summary>
    CsvEmpty,

    /// <summary>The <c>.csv</c> file's encoding is not UTF-8 or UTF-16 with a byte order mark.</summary>
    CsvUnsupportedEncoding,

    /// <summary>The <c>.csv</c> file has a quoted value that is never closed.</summary>
    CsvUnterminatedQuote,

    /// <summary>The changes need something the <c>.xlsx</c> writer does not do yet.</summary>
    SaveNotSupported,

    /// <summary>The document cannot be saved as it is (no sheet, a sheet without entries, two sheets with the same name).</summary>
    InvalidDocument,

    /// <summary>The <c>.xlsx</c> file the document was read from cannot be read again to save it.</summary>
    SourceDamaged,
}
