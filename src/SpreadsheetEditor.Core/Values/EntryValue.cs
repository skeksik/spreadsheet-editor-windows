using System.Globalization;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Values;

/// <summary>The value to write for an entry (docs/spec.md §9.2, §9.3, §9.16).</summary>
public static class EntryValue
{
    /// <summary>
    /// The value to write for an entry's cell, or null to clear it.
    /// <list type="bullet">
    /// <item>While the text equals the original value's display text, the original value, never parsed again.</item>
    /// <item>Empty text: null, the cell is cleared.</item>
    /// <item>Otherwise the text parsed for the culture (<see cref="CellText.Parse"/>). A date keeps the original number format when that is
    /// a date or time format, and otherwise gets the default for what was typed (<see cref="NumberFormats"/>). A number or boolean typed
    /// into a cell with a date or time format gets <see cref="NumberFormat.General"/>, so that it is not shown as a date. Other values
    /// carry no format: the writer leaves the cell's number format as it is.</item>
    /// </list>
    /// </summary>
    public static CellValue? Resolve(Entry entry, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(entry);
        ArgumentNullException.ThrowIfNull(culture);
        CellValue? original = entry.Original;
        if (original is not null && string.Equals(entry.Text, original.DisplayText, StringComparison.Ordinal))
        {
            return original;
        }

        if (entry.Text.Length == 0)
        {
            return null;
        }

        CellValue parsed = CellText.Parse(entry.Text, culture);
        bool originalIsDate = original is not null
            && (original.Kind == CellValueKind.Date || NumberFormats.IsDateOrTime(original.Format));
        return parsed.Kind switch
        {
            CellValueKind.Date => parsed with
            {
                Format = NumberFormats.IsDateOrTime(original?.Format) ? original!.Format : DefaultFormat(parsed.Date.GetValueOrDefault()),
            },
            CellValueKind.Number or CellValueKind.Boolean when originalIsDate => parsed with { Format = NumberFormat.General },
            _ => parsed,
        };
    }

    private static NumberFormat DefaultFormat(DateTime value)
    {
        TimeSpan time = value.TimeOfDay;
        if (DateSerial.IsTimeOnly(value))
        {
            return time.Seconds != 0 ? NumberFormats.LongTime : NumberFormats.ShortTime;
        }

        return time == TimeSpan.Zero ? NumberFormats.ShortDate : NumberFormats.DateAndTime;
    }
}
