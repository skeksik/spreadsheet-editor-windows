using System.Globalization;
using SpreadsheetEditor.Core.Values;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>What the sheets of one save share: the culture values were shown in, the workbook's tables, and what the save changed.</summary>
internal sealed class WriteContext(CultureInfo culture, DateSystem dateSystem, SharedStringWriter strings, StyleWriter styles)
{
    public CultureInfo Culture { get; } = culture;

    public DateSystem DateSystem { get; } = dateSystem;

    public SharedStringWriter Strings { get; } = strings;

    public StyleWriter Styles { get; } = styles;

    /// <summary>True once a cell value was written or cleared: formulas must be calculated again when the file is opened.</summary>
    public bool ValueChanged { get; set; }

    /// <summary>True once a cell that held a formula was changed or cleared, or a worksheet was deleted: the calculation chain is stale.</summary>
    public bool FormulaChanged { get; set; }
}
