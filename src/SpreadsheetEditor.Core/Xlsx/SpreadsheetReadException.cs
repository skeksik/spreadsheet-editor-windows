namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>A spreadsheet could not be read. <see cref="Exception.Message"/> is English text the app may show as the details.</summary>
public sealed class SpreadsheetReadException : Exception
{
    public SpreadsheetReadException()
    {
    }

    public SpreadsheetReadException(string message)
        : base(message)
    {
    }

    public SpreadsheetReadException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    public SpreadsheetReadException(SpreadsheetReadError reason, string message, Exception? innerException = null)
        : base(message, innerException)
    {
        Reason = reason;
    }

    /// <summary>Why the file could not be read.</summary>
    public SpreadsheetReadError Reason { get; }
}
