namespace SpreadsheetEditor.Core.Formulas;

/// <summary>
/// A piece of formula text: its kind and where it lies. A <see cref="FormulaTokenKind.Reference"/> also carries its parsed area, the
/// length of its sheet prefix (including the <c>!</c>; zero when it has none), the sheet name with quotes removed, and whether the
/// prefix names several sheets (<c>Sheet1:Sheet3!</c>) or another workbook (<c>[1]Sheet1!</c>), and whether it names another workbook.
/// An <see cref="FormulaTokenKind.Other"/> token that is a sheet prefix with no A1 area after it (<c>Sheet1!</c> in
/// <c>Sheet1!MyName</c>) carries the same prefix facts, with its whole length as the prefix.
/// </summary>
internal readonly record struct FormulaToken(FormulaTokenKind Kind, int Start, int Length)
{
    public A1Area Area { get; init; }

    public int PrefixLength { get; init; }

    public string? SheetName { get; init; }

    public bool IsMultiSheetOrExternal { get; init; }

    public bool IsExternal { get; init; }

    public int End => Start + Length;
}
