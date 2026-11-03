namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// A shared formula group as read: its index (the <c>f</c> element's <c>si</c>), its master cell, the master's <c>ref</c> and the
/// master's formula source, from which every other cell of the group takes its own formula (ECMA-376 Part 1 §18.3.1.40).
/// </summary>
/// <param name="Index">The group's <c>si</c>.</param>
/// <param name="MasterRow">The master cell's 1-based row.</param>
/// <param name="MasterColumn">The master cell's 1-based column.</param>
/// <param name="Range">The master's <c>ref</c> attribute as written (the range the group covers), or null when it has none.</param>
/// <param name="Source">The master's formula source without the leading <c>=</c>.</param>
internal sealed record SharedFormulaGroup(uint Index, int MasterRow, int MasterColumn, string? Range, string Source);
