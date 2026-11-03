using System.Globalization;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Values;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Builds a <see cref="CellValue"/> from a cell as stored in the worksheet: the reader's only place that constructs values. Every value
/// keeps the cell's number format as read (the writer relies on it, for example to keep edits in a Text cell as text); dates are told
/// from numbers by that format (<see cref="NumberFormats.Classify"/>) and converted in the workbook's date system
/// (<see cref="DateSerial"/>); the display text is <see cref="CellText.Format"/> (docs/spec.md §4.2, §9.2–9.3).
/// </summary>
internal static class XlsxCellValues
{
    /// <summary>The value of a cell, or null when it is empty (no value and no formula).</summary>
    /// <param name="type">The cell's <c>t</c> attribute (<c>s</c>, <c>inlineStr</c>, <c>str</c>, <c>b</c>, <c>e</c>, <c>d</c>, <c>n</c>), or null for a number.</param>
    /// <param name="raw">The text of the cell's <c>v</c> element, or null.</param>
    /// <param name="text">The resolved string of a shared or inline string cell, or null.</param>
    /// <param name="formula">The formula source without <c>=</c>, or null.</param>
    /// <param name="format">The cell style's number format.</param>
    /// <param name="system">The workbook's date system.</param>
    /// <param name="culture">The UI culture the display text is for.</param>
    public static CellValue? Create(
        string? type, string? raw, string? text, string? formula, NumberFormat format, DateSystem system, CultureInfo culture)
    {
        var value = Decode(type, raw, text, format, system);
        if (value is null)
        {
            if (formula is null)
            {
                return null;
            }
            value = new CellValue(CellValueKind.Text, "", Format: format);
        }
        value = value with { Formula = formula };
        return value with { DisplayText = CellText.Format(value, culture) };
    }

    private static CellValue? Decode(string? type, string? raw, string? text, NumberFormat format, DateSystem system)
    {
        switch (type)
        {
            case "s":
            case "inlineStr":
                return string.IsNullOrEmpty(text) ? null : new CellValue(CellValueKind.Text, text, Format: format);
            case "str":
                return string.IsNullOrEmpty(raw) ? null : new CellValue(CellValueKind.Text, raw, Format: format);
            case "e":
                return string.IsNullOrEmpty(raw) ? null : new CellValue(CellValueKind.Error, raw, Format: format);
            case "b":
                if (string.IsNullOrEmpty(raw))
                {
                    return null;
                }
                var flag = raw.Trim() is "1" or "true";
                return new CellValue(CellValueKind.Boolean, "", Boolean: flag, Format: format);
            case "d":
                // An ISO 8601 date and time (not verified against a file a spreadsheet program wrote).
                if (string.IsNullOrEmpty(raw))
                {
                    return null;
                }
                if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var iso))
                {
                    double? serial = DateSerial.TryFromDateTime(iso, system, out var s) ? s : null;
                    return new CellValue(CellValueKind.Date, "", Number: serial, Date: iso, Format: format);
                }
                return new CellValue(CellValueKind.Text, raw, Format: format);
            default:
                if (string.IsNullOrEmpty(raw))
                {
                    return null;
                }
                if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                {
                    return new CellValue(CellValueKind.Text, raw, Format: format);
                }
                if (NumberFormats.IsDateOrTime(format) && DateSerial.TryToDateTime(number, system, out var date))
                {
                    return new CellValue(CellValueKind.Date, "", Number: number, Date: date, Format: format);
                }
                return new CellValue(CellValueKind.Number, "", Number: number, Format: format);
        }
    }
}
