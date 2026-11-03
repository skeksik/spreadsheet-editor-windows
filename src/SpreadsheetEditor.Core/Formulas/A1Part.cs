namespace SpreadsheetEditor.Core.Formulas;

/// <summary>
/// One end of an A1 reference: a row, a column or both, each relative or absolute (<c>$</c>). A zero row means a whole-column part
/// (<c>A</c> in <c>A:C</c>); a zero column means a whole-row part (<c>3</c> in <c>1:3</c>).
/// </summary>
internal readonly record struct A1Part(int Row, bool RowAbsolute, int Column, bool ColumnAbsolute);
