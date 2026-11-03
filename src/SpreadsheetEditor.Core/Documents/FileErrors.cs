using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Documents;

/// <summary>Sorts the errors of opening and saving into the problems the app explains (docs/spec.md §10.17).</summary>
public static class FileErrors
{
    // HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) and HRESULT_FROM_WIN32(ERROR_LOCK_VIOLATION). .NET gives an IOException this HResult
    // for the Win32 error: dotnet/runtime, src/libraries/Common/src/System/IO/Win32Marshal.cs (main, read 2026-11-03), case
    // ERROR_SHARING_VIOLATION and the default case both pass MakeHRFromErrorCode(errorCode), which is "((int)0x80070000) | errorCode".
    private const int SharingViolation = unchecked((int)0x80070020);
    private const int LockViolation = unchecked((int)0x80070021);

    /// <summary>
    /// The problem behind <paramref name="error"/>, from the reader's and writer's reasons; <see cref="FileProblem.Unknown"/> for any other
    /// exception. A read or write that failed because another program holds the file (a sharing or lock violation anywhere in the inner
    /// exceptions) is <see cref="FileProblem.InUse"/>. Not every such case can be told apart: replacing a file another program has open
    /// may fail as access denied (seen in the owner's check, docs/notes/app-document-check.md), which stays <see cref="FileProblem.AccessDenied"/>.
    /// </summary>
    public static FileProblem Classify(Exception error)
    {
        ArgumentNullException.ThrowIfNull(error);
        if (error is SpreadsheetReadException or SpreadsheetWriteException && IsInUse(error))
        {
            return FileProblem.InUse;
        }
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

    private static bool IsInUse(Exception error)
    {
        for (var inner = error.InnerException; inner is not null; inner = inner.InnerException)
        {
            if (inner is IOException { HResult: SharingViolation or LockViolation })
            {
                return true;
            }
        }
        return false;
    }
}
