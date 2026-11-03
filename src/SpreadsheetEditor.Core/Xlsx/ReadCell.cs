using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>A non-empty worksheet cell as read: where it is, its value, and the plain string it holds when it is a string cell.</summary>
/// <param name="Row">The 1-based row.</param>
/// <param name="Column">The 1-based column.</param>
/// <param name="Value">The typed value and its display text.</param>
/// <param name="Text">The cell's string (shared, inline or a formula's string result), or null for a number, boolean, date or error.</param>
/// <param name="Formula">The formula source without the leading <c>=</c>, or null.</param>
internal sealed record ReadCell(int Row, int Column, CellValue Value, string? Text, string? Formula)
{
    /// <summary>
    /// Blank for the table-name test (docs/spec.md §4.1 step 4): a string of only white space. A formula is never blank, nor is a number,
    /// boolean, date or error.
    /// </summary>
    public bool IsBlank => Formula is null && Text is not null && string.IsNullOrWhiteSpace(Text);

    /// <summary>The string the cell holds when it is a plain string cell without a formula, else null.</summary>
    public string? PlainText => Formula is null ? Text : null;
}
