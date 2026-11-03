using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Values;

/// <summary>Recognises date and time number formats, built in or custom, and names the default formats for typed dates.</summary>
/// <remarks>
/// Built-in ids follow ECMA-376 Part 1 §18.8.30 (numFmt), as quoted in Microsoft's Open XML SDK reference for
/// <c>DocumentFormat.OpenXml.Spreadsheet.NumberingFormat</c> (ISO/IEC 29500-1 text, read 2026-11-03): "Following is a listing of number
/// formats whose formatCode value is implied rather than explicitly saved in the file." For all languages: 14 <c>mm-dd-yy</c>,
/// 15 <c>d-mmm-yy</c>, 16 <c>d-mmm</c>, 17 <c>mmm-yy</c>, 18 <c>h:mm AM/PM</c>, 19 <c>h:mm:ss AM/PM</c>, 20 <c>h:mm</c>, 21 <c>h:mm:ss</c>,
/// 22 <c>m/d/yy h:mm</c>, 45 <c>mm:ss</c>, 46 <c>[h]:mm:ss</c>, 47 <c>mmss.0</c>. For zh-tw, zh-cn, ja-jp and ko-kr, ids 27-36 and 50-58
/// are dates or times whose code depends on the language (for example 27 <c>[$-411]ge.m.d</c> in ja-jp, 32 <c>h"時"mm"分"</c>); ids 32
/// and 33 are times in all four, and 34, 35 and 52-56 are times in zh-tw or zh-cn but dates in ja-jp and ko-kr, so they are reported as
/// dates. For th-th: 71-74 and 81 dates, 75, 76 and 78-80 times, 77 date and time. "Ids not specified in the listing, such as 5, 6, 7,
/// and 8, shall follow the number format specified by the formatCode attribute."
/// </remarks>
public static class NumberFormats
{
    /// <summary>The default format for a typed date: built-in 14, a short date.</summary>
    public static NumberFormat ShortDate { get; } = new(14);

    /// <summary>The default format for a typed date and time: built-in 22, <c>m/d/yy h:mm</c>.</summary>
    public static NumberFormat DateAndTime { get; } = new(22);

    /// <summary>The default format for a typed time without seconds: built-in 20, <c>h:mm</c>.</summary>
    public static NumberFormat ShortTime { get; } = new(20);

    /// <summary>The default format for a typed time with seconds: built-in 21, <c>h:mm:ss</c>.</summary>
    public static NumberFormat LongTime { get; } = new(21);

    /// <summary>Whether a number in this format is a date or a time (any kind but <see cref="DateTimeFormatKind.None"/>).</summary>
    public static bool IsDateOrTime(NumberFormat? format) => Classify(format) != DateTimeFormatKind.None;

    /// <summary>
    /// Classifies a number format. A format code, when present, decides; otherwise the built-in id does. Null is <c>General</c>.
    /// </summary>
    public static DateTimeFormatKind Classify(NumberFormat? format)
    {
        if (format is null)
        {
            return DateTimeFormatKind.None;
        }

        return format.Code is { } code ? ClassifyCode(code) : ClassifyBuiltIn(format.Id);
    }

    /// <summary>Classifies a built-in number format id.</summary>
    public static DateTimeFormatKind ClassifyBuiltIn(int id) => id switch
    {
        >= 14 and <= 17 => DateTimeFormatKind.Date,
        >= 18 and <= 21 => DateTimeFormatKind.Time,
        22 => DateTimeFormatKind.DateAndTime,
        >= 45 and <= 47 => DateTimeFormatKind.Time,
        32 or 33 => DateTimeFormatKind.Time,
        >= 27 and <= 36 => DateTimeFormatKind.Date,
        >= 50 and <= 58 => DateTimeFormatKind.Date,
        >= 71 and <= 74 => DateTimeFormatKind.Date,
        75 or 76 or (>= 78 and <= 80) => DateTimeFormatKind.Time,
        77 => DateTimeFormatKind.DateAndTime,
        81 => DateTimeFormatKind.Date,
        _ => DateTimeFormatKind.None,
    };

    /// <summary>
    /// Classifies a format code by its first section. Quoted text, escaped characters (<c>\x</c>, <c>_x</c>, <c>*x</c>) and bracketed
    /// colours, conditions and locales are ignored; <c>[h]</c>, <c>[m]</c> and <c>[s]</c> are elapsed times. <c>y</c>, <c>d</c> and
    /// <c>e</c> (the era year) are date parts; <c>h</c>, <c>s</c>, <c>AM/PM</c> and <c>A/P</c> are time parts; <c>m</c> is minutes
    /// directly after an hour or directly before seconds, and the month otherwise.
    /// </summary>
    public static DateTimeFormatKind ClassifyCode(string code)
    {
        ArgumentNullException.ThrowIfNull(code);
        bool date = false;
        bool time = false;

        // The date and time letters in order ('y', 'd', 'h', 'm', 's'), to tell minutes from months.
        var parts = new List<char>();
        var monthOrMinute = new List<int>();

        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            switch (c)
            {
                case ';':
                    i = code.Length;
                    continue;
                case '"':
                    int close = code.IndexOf('"', i + 1);
                    i = close < 0 ? code.Length : close;
                    continue;
                case '\\' or '_' or '*':
                    i++;
                    continue;
                case '[':
                    int end = code.IndexOf(']', i + 1);
                    string inside = end < 0 ? code[(i + 1)..] : code[(i + 1)..end];
                    if (IsElapsed(inside))
                    {
                        time = true;
                        parts.Add(char.ToLowerInvariant(inside[0]));
                    }

                    i = end < 0 ? code.Length : end;
                    continue;
            }

            if (StartsWithAt(code, i, "General"))
            {
                i += "General".Length - 1;
                continue;
            }

            if (StartsWithAt(code, i, "AM/PM"))
            {
                time = true;
                i += "AM/PM".Length - 1;
                continue;
            }

            if (StartsWithAt(code, i, "A/P"))
            {
                time = true;
                i += "A/P".Length - 1;
                continue;
            }

            switch (char.ToLowerInvariant(c))
            {
                case 'e' when i + 1 < code.Length && code[i + 1] is '+' or '-':
                    // Scientific notation: 0.00E+00.
                    i++;
                    break;
                case 'y' or 'e':
                    date = true;
                    parts.Add('y');
                    break;
                case 'd':
                    date = true;
                    parts.Add('d');
                    break;
                case 'h':
                    time = true;
                    parts.Add('h');
                    break;
                case 's':
                    time = true;
                    parts.Add('s');
                    break;
                case 'm':
                    monthOrMinute.Add(parts.Count);
                    parts.Add('m');
                    break;
            }

            // Skip the rest of a run of the same letter (yyyy, mmm, ss).
            while (i + 1 < code.Length && char.ToLowerInvariant(code[i + 1]) == char.ToLowerInvariant(c) && char.IsLetter(c))
            {
                i++;
            }
        }

        foreach (int index in monthOrMinute)
        {
            bool afterHour = index > 0 && parts[index - 1] == 'h';
            bool beforeSecond = index + 1 < parts.Count && parts[index + 1] == 's';
            if (afterHour || beforeSecond)
            {
                time = true;
            }
            else
            {
                date = true;
            }
        }

        return (date, time) switch
        {
            (true, true) => DateTimeFormatKind.DateAndTime,
            (true, false) => DateTimeFormatKind.Date,
            (false, true) => DateTimeFormatKind.Time,
            _ => DateTimeFormatKind.None,
        };
    }

    private static bool IsElapsed(string inside)
    {
        if (inside.Length == 0)
        {
            return false;
        }

        char first = char.ToLowerInvariant(inside[0]);
        if (first is not ('h' or 'm' or 's'))
        {
            return false;
        }

        foreach (char c in inside)
        {
            if (char.ToLowerInvariant(c) != first)
            {
                return false;
            }
        }

        return true;
    }

    private static bool StartsWithAt(string text, int index, string value) =>
        string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0 && text.Length - index >= value.Length;
}
