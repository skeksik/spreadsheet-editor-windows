using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Documents;

/// <summary>A file that was read: its workbook, for <see cref="Editing.WorkbookEditor.Load"/>, and where it came from.</summary>
public sealed record OpenedDocument(Workbook Workbook, DocumentFile File);
