namespace SpreadsheetEditor.Core.Model;

/// <summary>
/// An open spreadsheet: the sheets that hold a table, in file order (docs/spec.md). The core is not thread-safe: a workbook is read or
/// written on one thread at a time, and edited on the UI thread only (ADR 0002).
/// </summary>
public sealed class Workbook
{
    public List<Sheet> Sheets { get; } = [];

    /// <summary>The file it was read from, or null for a new workbook.</summary>
    public string? SourcePath { get; set; }
}
