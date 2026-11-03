using SpreadsheetEditor.Core.Formulas;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// What one save does to every formula's references: the sheets renamed (docs/spec.md §9.30) and the rows each sheet inserts or deletes
/// (§9.20), in that order, so that row shifts name sheets by their new names.
/// </summary>
internal sealed class ReferenceTransform
{
    private readonly List<(string Old, string New)> renames = [];
    private readonly Dictionary<string, List<RowShift>> shifts = new(SheetNames.Comparer);

    /// <summary>True when the transform changes nothing.</summary>
    public bool IsEmpty => renames.Count == 0 && shifts.Count == 0;

    public void AddRename(string oldName, string newName)
    {
        if (!string.Equals(oldName, newName, StringComparison.Ordinal))
        {
            renames.Add((oldName, newName));
        }
    }

    /// <summary>The renamed sheets, old name to new.</summary>
    public IReadOnlyList<(string Old, string New)> Renames => renames;

    public void AddShifts(string sheetName, IReadOnlyList<RowShift> sheetShifts)
    {
        if (sheetShifts.Count > 0)
        {
            shifts[sheetName] = [.. sheetShifts];
        }
    }

    /// <summary>
    /// The formula with renamed sheets' prefixes rewritten and references to shifted rows moved.
    /// </summary>
    /// <param name="formula">The formula text, with or without its leading <c>=</c>.</param>
    /// <param name="formulaSheet">The (new) name of the sheet the formula lives on, which unqualified references point at; null for a
    /// workbook-level name or a chart.</param>
    public string Apply(string formula, string? formulaSheet)
    {
        var text = Rename(formula);
        foreach (var (_, list) in shifts)
        {
            foreach (var shift in list)
            {
                text = FormulaReferences.ShiftRows(text, formulaSheet, shift);
            }
        }

        return text;
    }

    /// <summary>
    /// The formula with renamed sheets' prefixes rewritten. With more than one rename, every old name is first changed to a placeholder,
    /// so that swapping two names does not rename both into one.
    /// </summary>
    public string Rename(string formula)
    {
        if (renames.Count == 1)
        {
            return FormulaReferences.RenameSheet(formula, renames[0].Old, renames[0].New);
        }

        var text = formula;
        for (var i = 0; i < renames.Count; i++)
        {
            text = FormulaReferences.RenameSheet(text, renames[i].Old, Placeholder(i));
        }

        for (var i = 0; i < renames.Count; i++)
        {
            text = FormulaReferences.RenameSheet(text, Placeholder(i), renames[i].New);
        }

        return text;
    }

    /// <summary>A sheet name's new name, or the name itself when it was not renamed.</summary>
    public string RenameSheetName(string name)
    {
        foreach (var (old, renamed) in renames)
        {
            if (SheetNames.Comparer.Equals(old, name))
            {
                return renamed;
            }
        }

        return name;
    }

    // A name made with a private-use character, which no sheet name in a real workbook is expected to hold.
    private static string Placeholder(int index) => "\uE000rename" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
