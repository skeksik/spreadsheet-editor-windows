namespace SpreadsheetEditor.Core.Documents;

/// <summary>What a save did.</summary>
/// <param name="File">The document's file after the save: the path and format Save writes next time.</param>
/// <param name="AllSheetsWritten">
/// False when a <c>.csv</c> save wrote one sheet of several: the document is then modified (docs/spec.md §9.34), so the app calls
/// <see cref="Editing.WorkbookEditor.MarkModified"/> instead of <see cref="Editing.WorkbookEditor.MarkSaved"/>.
/// </param>
public sealed record SaveResult(DocumentFile File, bool AllSheetsWritten);
