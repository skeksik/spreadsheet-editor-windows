using System.Globalization;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Builds a <see cref="CellValue"/> from a cell as stored in the worksheet. This is the reader's only place that constructs values:
/// feat/cell-values replaces the interim date test, date conversion and display text here with its own (<c>NumberFormats</c>,
/// <c>DateSerial</c>, <c>CellText.Format</c>) and adds the formula and number format to the value.
/// </summary>
internal static class XlsxCellValues
{
    /// <summary>
    /// The value of a cell, or null when it is empty (no value and no formula).
    /// </summary>
    /// <param name="type">The cell's <c>t</c> attribute (<c>s</c>, <c>inlineStr</c>, <c>str</c>, <c>b</c>, <c>e</c>, <c>d</c>, <c>n</c>), or null for a number.</param>
    /// <param name="raw">The text of the cell's <c>v</c> element, or null.</param>
    /// <param name="text">The resolved string of a shared or inline string cell, or null.</param>
    /// <param name="formula">The formula source without <c>=</c>, or null.</param>
    /// <param name="formatId">The cell style's <c>numFmtId</c>.</param>
    /// <param name="formatCode">The format code of a custom <paramref name="formatId"/>, or null.</param>
    /// <param name="date1904">True when the workbook uses the 1904 date system.</param>
    /// <param name="culture">The UI culture the display text is for.</param>
    public static CellValue? Create(
        string? type, string? raw, string? text, string? formula, int formatId, string? formatCode, bool date1904, CultureInfo culture)
    {
        var value = Decode(type, raw, text, formatId, formatCode, date1904, culture);
        if (formula is not null)
        {
            // Until CellValue carries the formula, the display text is the formula source (docs/spec.md §4.2, §9.3).
            return (value ?? new CellValue(CellValueKind.Text, "")) with { DisplayText = "=" + formula };
        }
        return value;
    }

    private static CellValue? Decode(
        string? type, string? raw, string? text, int formatId, string? formatCode, bool date1904, CultureInfo culture)
    {
        switch (type)
        {
            case "s":
            case "inlineStr":
                return string.IsNullOrEmpty(text) ? null : new CellValue(CellValueKind.Text, text);
            case "str":
                return string.IsNullOrEmpty(raw) ? null : new CellValue(CellValueKind.Text, raw);
            case "e":
                // Errors (#N/A, #DIV/0!) are shown as their text.
                return string.IsNullOrEmpty(raw) ? null : new CellValue(CellValueKind.Text, raw);
            case "b":
                if (string.IsNullOrEmpty(raw))
                {
                    return null;
                }
                var flag = raw.Trim() is "1" or "true";
                return new CellValue(CellValueKind.Boolean, flag ? "True" : "False", Boolean: flag);
            case "d":
                if (string.IsNullOrEmpty(raw))
                {
                    return null;
                }
                if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var iso))
                {
                    return new CellValue(CellValueKind.Date, DateText(iso, culture), Number: iso.ToOADate(), Date: iso);
                }
                return new CellValue(CellValueKind.Text, raw);
            default:
                if (string.IsNullOrEmpty(raw))
                {
                    return null;
                }
                if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                {
                    return new CellValue(CellValueKind.Text, raw);
                }
                if (IsDateFormat(formatId, formatCode) && TrySerialToDate(number, date1904, out var date))
                {
                    return new CellValue(CellValueKind.Date, DateText(date, culture), Number: number, Date: date);
                }
                return new CellValue(CellValueKind.Number, number.ToString(culture), Number: number);
        }
    }

    // Interim: feat/cell-values formats dates in the Windows short date format of the UI language.
    private static string DateText(DateTime date, CultureInfo culture) =>
        date.TimeOfDay == TimeSpan.Zero ? date.ToString("d", culture) : date.ToString("g", culture);

    /// <summary>
    /// Interim date test. Built-in ids are the date and time formats listed in ISO/IEC 29500-1 §18.8.30 (quoted in the Open XML SDK's
    /// NumberingFormat reference, read 2026-11-03): 14–22, 45–47, and the East Asian 27–36 and 50–58. A custom code is a date when it
    /// has a date or time letter outside quoted text, escapes and brackets other than elapsed time.
    /// </summary>
    internal static bool IsDateFormat(int formatId, string? formatCode)
    {
        if (formatCode is null)
        {
            return formatId is (>= 14 and <= 22) or (>= 27 and <= 36) or (>= 45 and <= 47) or (>= 50 and <= 58);
        }
        var inQuotes = false;
        for (var i = 0; i < formatCode.Length; i++)
        {
            var c = formatCode[i];
            if (inQuotes)
            {
                inQuotes = c != '"';
                continue;
            }
            switch (c)
            {
                case '"':
                    inQuotes = true;
                    break;
                case '\\':
                case '_':
                case '*':
                    i++;
                    break;
                case '[':
                    var end = formatCode.IndexOf(']', i);
                    if (end < 0)
                    {
                        return false;
                    }
                    var inside = formatCode.AsSpan(i + 1, end - i - 1);
                    if (inside.Length > 0 && inside.Trim("hHmMsS").IsEmpty)
                    {
                        return true;
                    }
                    i = end;
                    break;
                case 'd' or 'D' or 'm' or 'M' or 'y' or 'Y' or 'h' or 'H' or 's' or 'S':
                    return true;
                case 'G' or 'g' when formatCode.AsSpan(i).StartsWith("General", StringComparison.OrdinalIgnoreCase):
                    i += "General".Length - 1;
                    break;
            }
        }
        return false;
    }

    // Interim: the 1900 system counts the nonexistent 1900-02-29 (serial 60); below it, serials are one day off the OLE date.
    private static bool TrySerialToDate(double serial, bool date1904, out DateTime date)
    {
        date = default;
        var oa = date1904 ? serial + 1462 : serial < 60 ? serial + 1 : serial;
        if (serial < 0 || oa >= 2958466)
        {
            return false;
        }
        date = DateTime.FromOADate(oa);
        return true;
    }
}
