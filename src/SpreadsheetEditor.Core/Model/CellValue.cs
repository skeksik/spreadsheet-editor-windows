namespace SpreadsheetEditor.Core.Model;

public enum CellValueKind
{
    Text,
    Number,
    Boolean,
    Date,

    /// <summary>An error value such as <c>#DIV/0!</c>, kept as its text.</summary>
    Error,
}

/// <summary>A typed cell value as read from a file, with the text the app showed for it.</summary>
/// <param name="Kind">The type of the value; for a formula cell, the type of its cached result.</param>
/// <param name="DisplayText">
/// The text shown for the value (<see cref="Values.CellText.Format"/>). For <see cref="CellValueKind.Text"/> and
/// <see cref="CellValueKind.Error"/> values it is also the value itself.
/// </param>
/// <param name="Number">A number's value; for a date read from a file, its serial as read (null for a date parsed from typed text).</param>
/// <param name="Boolean">A boolean's value.</param>
/// <param name="Date">A date's value: the date and time of the serial, or a time on <see cref="DateTime.MinValue"/>'s day for a time of
/// day only (<see cref="Values.DateSerial"/>).</param>
/// <param name="Formula">A formula cell's source without the leading <c>=</c>, or null when the cell holds no formula.</param>
/// <param name="Format">The cell's number format as read, or for a value parsed from typed text the format to write (null: leave the
/// cell's number format as it is).</param>
public sealed record CellValue(
    CellValueKind Kind,
    string DisplayText,
    double? Number = null,
    bool? Boolean = null,
    DateTime? Date = null,
    string? Formula = null,
    NumberFormat? Format = null);
