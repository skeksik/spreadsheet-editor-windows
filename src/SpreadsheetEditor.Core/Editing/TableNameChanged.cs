using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>The sheet's <see cref="Sheet.TableName"/> changed.</summary>
public sealed record TableNameChanged(Sheet Sheet) : WorkbookChange(Sheet);
