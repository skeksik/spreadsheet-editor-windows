namespace SpreadsheetEditor.Core.Csv;

/// <summary>A .csv file could not be read. The app shows a message for <see cref="Reason"/>; <see cref="Exception.Message"/> is English.</summary>
public sealed class CsvReadException : Exception
{
    public CsvReadException()
    {
    }

    public CsvReadException(string message)
        : base(message)
    {
    }

    public CsvReadException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    public CsvReadException(CsvReadFailure reason, int lineNumber, string message, Exception? innerException = null)
        : base(message, innerException)
    {
        Reason = reason;
        LineNumber = lineNumber;
    }

    public CsvReadFailure Reason { get; }

    /// <summary>The 1-based line the problem was found on, or 0 when it has none.</summary>
    public int LineNumber { get; }
}
