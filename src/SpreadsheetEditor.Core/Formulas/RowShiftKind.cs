namespace SpreadsheetEditor.Core.Formulas;

/// <summary>Whether a <see cref="RowShift"/> inserts or deletes worksheet rows.</summary>
public enum RowShiftKind
{
    /// <summary>New rows are inserted; the rows from the first one down move down.</summary>
    Insert,

    /// <summary>Rows are deleted; the rows below them move up.</summary>
    Delete,
}
