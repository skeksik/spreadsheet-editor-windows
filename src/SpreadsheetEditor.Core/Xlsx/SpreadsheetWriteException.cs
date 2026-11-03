namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// A workbook could not be saved. <see cref="Exception.Message"/> is English text the app may show as the details. When it is thrown, the
/// target file is unchanged: nothing is written to it until the whole workbook has been written.
/// </summary>
public sealed class SpreadsheetWriteException : Exception
{
    public SpreadsheetWriteException()
    {
    }

    public SpreadsheetWriteException(string message)
        : base(message)
    {
    }

    public SpreadsheetWriteException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    public SpreadsheetWriteException(SpreadsheetWriteError reason, string message, Exception? innerException = null)
        : base(message, innerException)
    {
        Reason = reason;
    }

    /// <summary>Why the workbook could not be saved.</summary>
    public SpreadsheetWriteError Reason { get; }

    internal static SpreadsheetWriteException NotSupported(string message) =>
        new(SpreadsheetWriteError.NotSupported, message + " This is not supported yet.");

    internal static SpreadsheetWriteException Invalid(string message) => new(SpreadsheetWriteError.InvalidWorkbook, message);
}
