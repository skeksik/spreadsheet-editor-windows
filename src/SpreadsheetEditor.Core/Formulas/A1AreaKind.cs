namespace SpreadsheetEditor.Core.Formulas;

/// <summary>What an A1 reference covers.</summary>
internal enum A1AreaKind
{
    /// <summary>A cell (<c>B2</c>) or a range of cells (<c>B2:D9</c>).</summary>
    Cells,

    /// <summary>Whole rows (<c>1:3</c>).</summary>
    Rows,

    /// <summary>Whole columns (<c>A:C</c>).</summary>
    Columns,
}
