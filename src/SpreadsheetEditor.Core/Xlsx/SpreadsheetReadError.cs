namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Why a spreadsheet could not be read; the app shows a message for each reason and the exception's message as details.</summary>
public enum SpreadsheetReadError
{
    /// <summary>The file is a zip package but not a readable workbook: a part is missing or its XML is damaged.</summary>
    Corrupt,

    /// <summary>The file or its folder does not exist.</summary>
    NotFound,

    /// <summary>The file cannot be opened for reading (permissions).</summary>
    AccessDenied,

    /// <summary>The file could not be read, for example because another program has it locked.</summary>
    ReadFailed,

    /// <summary>The file is not an <c>.xlsx</c> package at all.</summary>
    NotAnXlsx,

    /// <summary>The file is a compound document: an older <c>.xls</c> workbook or an encrypted (password-protected) <c>.xlsx</c>.</summary>
    LegacyOrEncrypted,
}
