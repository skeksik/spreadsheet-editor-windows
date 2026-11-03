namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Why a workbook could not be saved; the app shows a message for each reason and the exception's message as details.</summary>
public enum SpreadsheetWriteError
{
    /// <summary>
    /// The change needs something the writer does not do yet, such as moving rows on a sheet with a pivot table, or laying out pictures in
    /// new cells.
    /// </summary>
    NotSupported,

    /// <summary>The workbook cannot be saved as it is: no sheet, a sheet without entries, or two worksheets with the same name.</summary>
    InvalidWorkbook,

    /// <summary>The <c>.xlsx</c> the workbook was read from cannot be opened again: a part is missing or its XML is damaged.</summary>
    SourceDamaged,

    /// <summary>The target file or its folder cannot be written (permissions, a read-only file).</summary>
    AccessDenied,

    /// <summary>The file could not be written, for example because the folder does not exist or another program has the file locked.</summary>
    WriteFailed,
}
