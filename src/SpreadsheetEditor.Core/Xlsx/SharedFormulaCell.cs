namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>A worksheet cell's membership in a shared formula group: the group's <c>si</c> and whether the cell is its master.</summary>
/// <param name="Index">The group's <c>si</c>.</param>
/// <param name="IsMaster">True for the cell that carries the group's formula text and <c>ref</c>.</param>
internal readonly record struct SharedFormulaCell(uint Index, bool IsMaster);
