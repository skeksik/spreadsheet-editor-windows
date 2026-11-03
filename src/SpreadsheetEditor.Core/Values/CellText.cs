using System.Globalization;
using System.Text;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Values;

/// <summary>The text of a cell value for a culture, and edited text parsed back into a value (docs/spec.md §4.2, §9.2-9.4).</summary>
public static class CellText
{
    /// <summary>The text of a boolean in every culture, as typed and as shown.</summary>
    public const string True = "TRUE";

    /// <inheritdoc cref="True"/>
    public const string False = "FALSE";

    /// <summary>
    /// Significant digits a typed number may have. Longer digit strings (card numbers, long codes) stay text, since a number would lose
    /// digits. Microsoft, "Excel specifications and limits" (read 2026-11-03): "Number precision | 15 digits".
    /// </summary>
    public const int MaxSignificantDigits = 15;

    /// <summary>
    /// The text shown for a value. A formula is <c>=</c> and its source. A number is the shortest text that reads back as the same number,
    /// in fixed-point notation, with the culture's decimal separator and no group separators; the number format is not applied. A boolean
    /// is <see cref="True"/> or <see cref="False"/>. A date is the culture's short date, followed by the time (short, or long when it has
    /// seconds) when it has a time part; a time of day only is the time. Text and errors are <see cref="CellValue.DisplayText"/>.
    /// </summary>
    public static string Format(CellValue value, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(value);
        ArgumentNullException.ThrowIfNull(culture);
        if (value.Formula is { } formula)
        {
            return "=" + formula;
        }

        return value.Kind switch
        {
            CellValueKind.Number when value.Number is { } number => FormatNumber(number, culture),
            CellValueKind.Boolean when value.Boolean is { } boolean => boolean ? True : False,
            CellValueKind.Date when value.Date is { } date => FormatDate(date, culture),
            _ => value.DisplayText,
        };
    }

    /// <summary>The text of a number (<see cref="Format"/>).</summary>
    public static string FormatNumber(double number, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(culture);
        if (double.IsNaN(number) || double.IsInfinity(number))
        {
            return number.ToString(culture);
        }

        if (number == 0)
        {
            return "0";
        }

        // Shortest round-trip digits. Microsoft, "Standard numeric format strings" (read 2026-11-03), default precision of "G" for
        // Double: "Smallest round-trippable number of digits to represent the number".
        string shortest = number.ToString(CultureInfo.InvariantCulture);
        bool negative = shortest[0] == '-';
        if (negative)
        {
            shortest = shortest[1..];
        }

        string fixedPoint = ToFixedPoint(shortest);
        NumberFormatInfo info = culture.NumberFormat;
        string text = fixedPoint.Replace(".", info.NumberDecimalSeparator, StringComparison.Ordinal);
        return negative ? info.NegativeSign + text : text;
    }

    /// <summary>The text of a date, a date and time, or a time of day only (<see cref="Format"/>).</summary>
    public static string FormatDate(DateTime date, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(culture);
        DateTimeFormatInfo info = culture.DateTimeFormat;
        TimeSpan time = date.TimeOfDay;
        string timePattern = time.Seconds != 0 || time.Milliseconds != 0 ? info.LongTimePattern : info.ShortTimePattern;
        if (DateSerial.IsTimeOnly(date))
        {
            return date.ToString(timePattern, culture);
        }

        string text = date.ToString(info.ShortDatePattern, culture);
        return time == TimeSpan.Zero ? text : text + " " + date.ToString(timePattern, culture);
    }

    /// <summary>
    /// Parses edited text for a culture. The result's <see cref="CellValue.DisplayText"/> is the text exactly as given.
    /// <list type="bullet">
    /// <item><c>=</c> followed by anything: a formula with that source and no cached value.</item>
    /// <item><see cref="True"/> or <see cref="False"/>, ignoring case: a boolean.</item>
    /// <item>A number in the culture's format: an optional negative sign, digits, and optionally the culture's decimal separator and
    /// digits. No spaces, group separators, plus sign, exponent, currency or percent; no leading zero before other digits ("007" is a
    /// code); at most <see cref="MaxSignificantDigits"/> significant digits.</item>
    /// <item>A date in the culture's short date pattern, optionally followed by a space and a time in its short or long time pattern, or
    /// such a time alone: a date (a time alone lies on <see cref="DateSerial.TimeOnlyDay"/>). Years from 1900 to 9999.</item>
    /// <item>Anything else: text, unchanged ("2024-01-15 meeting notes" stays text).</item>
    /// </list>
    /// </summary>
    public static CellValue Parse(string text, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(culture);

        if (text.Length > 1 && text[0] == '=')
        {
            return new CellValue(CellValueKind.Text, text, Formula: text[1..]);
        }

        if (string.Equals(text, True, StringComparison.OrdinalIgnoreCase))
        {
            return new CellValue(CellValueKind.Boolean, text, Boolean: true);
        }

        if (string.Equals(text, False, StringComparison.OrdinalIgnoreCase))
        {
            return new CellValue(CellValueKind.Boolean, text, Boolean: false);
        }

        if (TryParseNumber(text, culture, out double number))
        {
            return new CellValue(CellValueKind.Number, text, Number: number);
        }

        if (TryParseDate(text, culture, out DateTime date))
        {
            return new CellValue(CellValueKind.Date, text, Date: date);
        }

        return new CellValue(CellValueKind.Text, text);
    }

    /// <summary>Parses a number as <see cref="Parse"/> does.</summary>
    public static bool TryParseNumber(string text, CultureInfo culture, out double number)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(culture);
        number = 0;
        NumberFormatInfo info = culture.NumberFormat;

        int i = 0;
        if (info.NegativeSign.Length > 0 && text.StartsWith(info.NegativeSign, StringComparison.Ordinal))
        {
            i = info.NegativeSign.Length;
        }

        int integerStart = i;
        while (i < text.Length && char.IsAsciiDigit(text[i]))
        {
            i++;
        }

        int integerDigits = i - integerStart;
        if (integerDigits == 0 || (integerDigits > 1 && text[integerStart] == '0'))
        {
            return false;
        }

        int fractionDigits = 0;
        string separator = info.NumberDecimalSeparator;
        if (i < text.Length)
        {
            if (separator.Length == 0 || string.CompareOrdinal(text, i, separator, 0, separator.Length) != 0)
            {
                return false;
            }

            i += separator.Length;
            int fractionStart = i;
            while (i < text.Length && char.IsAsciiDigit(text[i]))
            {
                i++;
            }

            fractionDigits = i - fractionStart;
            if (fractionDigits == 0 || i != text.Length)
            {
                return false;
            }
        }

        string integer = text.Substring(integerStart, integerDigits);
        string fraction = text[(text.Length - fractionDigits)..];
        if ((integer + fraction).Trim('0').Length > MaxSignificantDigits)
        {
            return false;
        }

        string invariant = fractionDigits == 0 ? integer : integer + "." + fraction;
        number = double.Parse(invariant, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
        if (integerStart > 0 && number != 0)
        {
            number = -number;
        }

        return true;
    }

    /// <summary>Parses a date, a date and time, or a time as <see cref="Parse"/> does.</summary>
    public static bool TryParseDate(string text, CultureInfo culture, out DateTime date)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(culture);
        DateTimeFormatInfo info = culture.DateTimeFormat;
        string shortDate = NormaliseSpaces(info.ShortDatePattern);
        string shortTime = NormaliseSpaces(info.ShortTimePattern);
        string longTime = NormaliseSpaces(info.LongTimePattern);
        string[] dateFormats = [shortDate, shortDate + " " + shortTime, shortDate + " " + longTime];
        string[] timeFormats = [shortTime, longTime];
        string input = NormaliseSpaces(text);

        if (DateTime.TryParseExact(input, dateFormats, culture, DateTimeStyles.None, out date))
        {
            return date.Year >= 1900;
        }

        if (DateTime.TryParseExact(input, timeFormats, culture, DateTimeStyles.NoCurrentDateDefault, out date))
        {
            return DateSerial.IsTimeOnly(date);
        }

        return false;
    }

    // Recent ICU versions put a narrow no-break space (U+202F) or a no-break space (U+00A0) in time patterns, for example before AM/PM;
    // a typed ordinary space should match it (not verified on every ICU version or on Windows' NLS data).
    private static string NormaliseSpaces(string text) => text.Replace(' ', ' ').Replace(' ', ' ');

    private static string ToFixedPoint(string shortest)
    {
        int e = shortest.IndexOf('E', StringComparison.Ordinal);
        if (e < 0)
        {
            return shortest;
        }

        string mantissa = shortest[..e];
        int exponent = int.Parse(shortest.AsSpan(e + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        int point = mantissa.IndexOf('.', StringComparison.Ordinal);
        string digits = point < 0 ? mantissa : mantissa.Remove(point, 1);
        int integerDigits = (point < 0 ? mantissa.Length : point) + exponent;

        // Very large or very small numbers keep the exponent, as a fixed-point form would be too long to read.
        if (integerDigits > 21 || integerDigits < -20)
        {
            return shortest;
        }

        var builder = new StringBuilder();
        if (integerDigits <= 0)
        {
            builder.Append("0.").Append('0', -integerDigits).Append(digits);
        }
        else if (integerDigits >= digits.Length)
        {
            builder.Append(digits).Append('0', integerDigits - digits.Length);
        }
        else
        {
            builder.Append(digits, 0, integerDigits).Append('.').Append(digits, integerDigits, digits.Length - integerDigits);
        }

        return builder.ToString();
    }
}
