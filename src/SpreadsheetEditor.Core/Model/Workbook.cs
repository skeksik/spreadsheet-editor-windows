namespace SpreadsheetEditor.Core.Model;

/// <summary>
/// An open spreadsheet: its grid worksheets as sheets, in file order (docs/spec.md). The core is not thread-safe: a workbook is read or
/// written on one thread at a time, and edited on the UI thread only (ADR 0002).
/// </summary>
public sealed class Workbook
{
    public List<Sheet> Sheets { get; } = [];

    /// <summary>
    /// The names of the sheets the app does not show (chart, dialog and macro sheets), kept in the file as they are (docs/spec.md §9.12).
    /// Empty worksheets are not among them: they open as sheets (§9.10). A sheet the user adds or renames must not take one of them (§9.39).
    /// </summary>
    public List<string> SkippedSheetNames { get; } = [];

    /// <summary>The file it was read from, or null for a new workbook.</summary>
    public string? SourcePath { get; set; }

    /// <summary>The <c>.xlsx</c> package it was read from, kept for the writer; null for a new workbook or one read from CSV.</summary>
    internal Xlsx.XlsxSource? Source { get; set; }
}
