using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Documents;

/// <summary>Sorts the errors of opening and saving into the problems the app explains (docs/spec.md §10.17).</summary>
public static class FileErrors
{
    /// <summary>The problem behind <paramref name="error"/>, from the reader's and writer's reasons; <see cref="FileProblem.Unknown"/> for any other exception.</summary>
    public static FileProblem Classify(Exception error)
    {
        ArgumentNullException.ThrowIfNull(error);
        return error switch
        {
            SpreadsheetReadException read => read.Reason switch
            {
                SpreadsheetReadError.NotFound => FileProblem.NotFound,
                SpreadsheetReadError.AccessDenied => FileProblem.AccessDenied,
                SpreadsheetReadError.ReadFailed => FileProblem.ReadFailed,
                SpreadsheetReadError.NotAnXlsx => FileProblem.NotAWorkbook,
                SpreadsheetReadError.LegacyOrEncrypted => FileProblem.LegacyOrEncrypted,
                SpreadsheetReadError.Corrupt => FileProblem.Damaged,
                _ => FileProblem.Unknown,
            },
            CsvReadException csv => csv.Reason switch
            {
                CsvReadFailure.Empty => FileProblem.CsvEmpty,
                CsvReadFailure.UnsupportedEncoding => FileProblem.CsvUnsupportedEncoding,
                CsvReadFailure.UnterminatedQuote => FileProblem.CsvUnterminatedQuote,
                _ => FileProblem.Unknown,
            },
            SpreadsheetWriteException write => write.Reason switch
            {
                SpreadsheetWriteError.NotSupported => FileProblem.SaveNotSupported,
                SpreadsheetWriteError.InvalidWorkbook => FileProblem.InvalidDocument,
                SpreadsheetWriteError.SourceDamaged => FileProblem.SourceDamaged,
                SpreadsheetWriteError.AccessDenied => FileProblem.AccessDenied,
                SpreadsheetWriteError.WriteFailed => FileProblem.WriteFailed,
                _ => FileProblem.Unknown,
            },
            _ => FileProblem.Unknown,
        };
    }
}
