using System.Globalization;
using SpreadsheetEditor.Core.Values;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>What the sheets of one save share: the culture values were shown in, the workbook's tables, the references the save rewrites, and
/// what the save changed.</summary>
internal sealed class WriteContext(
    CultureInfo culture, DateSystem dateSystem, SharedStringWriter strings, StyleWriter styles, ReferenceTransform references)
{
    public CultureInfo Culture { get; } = culture;

    public DateSystem DateSystem { get; } = dateSystem;

    public SharedStringWriter Strings { get; } = strings;

    public StyleWriter Styles { get; } = styles;

    /// <summary>The save's sheet renames and row shifts, which a formula moved to another cell is rewritten with.</summary>
    public ReferenceTransform References { get; } = references;

    /// <summary>True once a cell value was written or cleared: formulas must be calculated again when the file is opened.</summary>
    public bool ValueChanged { get; set; }

    /// <summary>True once a cell that held a formula was changed or cleared, or a worksheet was deleted: the calculation chain is stale.</summary>
    public bool FormulaChanged { get; set; }
}
