namespace SpreadsheetEditor.Core.Model;

/// <summary>
/// The result of <see cref="SheetNames.Check"/>: the first rule a worksheet name breaks, and for
/// <see cref="SheetNameError.InvalidCharacter"/> the first character that breaks it.
/// </summary>
public readonly record struct SheetNameCheck(SheetNameError Error, char? Character = null)
{
    public static SheetNameCheck Valid { get; }

    public bool IsValid => Error == SheetNameError.None;
}
