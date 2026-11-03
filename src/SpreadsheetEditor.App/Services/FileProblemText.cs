using SpreadsheetEditor.Core.Documents;

namespace SpreadsheetEditor.App.Services;

/// <summary>The plain explanation of a <see cref="FileProblem"/> (docs/spec.md §10.17), from the resources <c>FileProblem_&lt;name&gt;</c>.</summary>
internal static class FileProblemText
{
    public static string For(FileProblem problem) => Strings.Get(problem switch
    {
        FileProblem.NotFound => "FileProblem_NotFound",
        FileProblem.AccessDenied => "FileProblem_AccessDenied",
        FileProblem.InUse => "FileProblem_InUse",
        FileProblem.ReadFailed => "FileProblem_ReadFailed",
        FileProblem.WriteFailed => "FileProblem_WriteFailed",
        FileProblem.NotAWorkbook => "FileProblem_NotAWorkbook",
        FileProblem.LegacyOrEncrypted => "FileProblem_LegacyOrEncrypted",
        FileProblem.Damaged => "FileProblem_Damaged",
        FileProblem.CsvEmpty => "FileProblem_CsvEmpty",
        FileProblem.CsvUnsupportedEncoding => "FileProblem_CsvUnsupportedEncoding",
        FileProblem.CsvUnterminatedQuote => "FileProblem_CsvUnterminatedQuote",
        FileProblem.SaveNotSupported => "FileProblem_SaveNotSupported",
        FileProblem.InvalidDocument => "FileProblem_InvalidDocument",
        FileProblem.SourceDamaged => "FileProblem_SourceDamaged",
        _ => "FileProblem_Unknown",
    });
}
