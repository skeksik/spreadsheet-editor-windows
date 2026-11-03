namespace SpreadsheetEditor.Core.Formulas;

/// <summary>
/// A piece of formula text: its kind and where it lies. A <see cref="FormulaTokenKind.Reference"/> also carries its parsed area, the
/// length of its sheet prefix (including the <c>!</c>; zero when it has none), the sheet name with quotes removed, and whether the
/// prefix names several sheets (<c>Sheet1:Sheet3!</c>) or another workbook (<c>[1]Sheet1!</c>).
/// </summary>
internal readonly record struct FormulaToken(FormulaTokenKind Kind, int Start, int Length)
{
    public A1Area Area { get; init; }

    public int PrefixLength { get; init; }

    public string? SheetName { get; init; }

    public bool IsMultiSheetOrExternal { get; init; }

    public int End => Start + Length;
}
