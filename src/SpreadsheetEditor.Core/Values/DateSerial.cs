namespace SpreadsheetEditor.Core.Values;

/// <summary>
/// Converts spreadsheet date serials (days since the date system's start, the time of day as the fraction) to and from <see cref="DateTime"/>.
/// </summary>
/// <remarks>
/// Sources, read 2026-11-03:
/// <list type="bullet">
/// <item>Microsoft, "DATE function": "January 1, 1900 is serial number 1"; "January 1, 2008 is serial number 39448".</item>
/// <item>Microsoft, "Date systems in Excel": "In the 1904 date system, dates are calculated by using January 1, 1904, as a starting
/// point."; "The difference between the two date systems is 1,462 days."; 5 July 2011 is 40729 (1900) and 39267 (1904).</item>
/// <item>Microsoft, "Excel incorrectly assumes that the year 1900 is a leap year": Excel keeps the fictitious 29 February 1900 for
/// compatibility with Lotus 1-2-3, so serial 60 is that day and every serial from 61 (1 March 1900) on is one more than the days since
/// 31 December 1899.</item>
/// <item>Microsoft, "Excel specifications and limits": earliest date "January 1, 1900 (January 1, 1904, if 1904 date system is used)",
/// latest date "December 31, 9999".</item>
/// </list>
/// ECMA-376 Part 1 §18.17.4 describes the same date systems; it was not read for this code (not verified against it).
/// <para>
/// A serial below 1 has no date part: it is a time of day only, and is represented by that time on <see cref="DateTime.MinValue"/>'s day
/// (1 January 0001) in both date systems. In the 1904 system serial 0 is also a real date, 1 January 1904, which this reads as midnight
/// only; the reader decides between a date and a time by the cell's number format (<see cref="NumberFormats.Classify"/>).
/// </para>
/// </remarks>
public static class DateSerial
{
    private const double MillisecondsPerDay = 86_400_000;

    private static readonly DateTime Base1900 = new(1899, 12, 31);
    private static readonly DateTime Base1900AfterLeapBug = new(1899, 12, 30);
    private static readonly DateTime FirstDayAfterLeapBug = new(1900, 3, 1);
    private static readonly DateTime Base1904 = new(1904, 1, 1);

    /// <summary>The day on which a time-of-day-only value lies.</summary>
    public static DateTime TimeOnlyDay => DateTime.MinValue;

    /// <summary>Whether a value is a time of day only: it lies on <see cref="TimeOnlyDay"/>.</summary>
    public static bool IsTimeOnly(DateTime value) => value.Date == TimeOnlyDay;

    /// <summary>
    /// Converts a serial to a date and time, rounded to the millisecond. False for a negative serial, a serial past 31 December 9999,
    /// and serial 60 in the 1900 system (29 February 1900, which does not exist).
    /// </summary>
    public static bool TryToDateTime(double serial, DateSystem system, out DateTime value)
    {
        value = default;
        if (double.IsNaN(serial) || serial < 0 || serial >= 3_000_000)
        {
            return false;
        }

        double days = Math.Floor(serial);
        long milliseconds = (long)Math.Round((serial - days) * MillisecondsPerDay, MidpointRounding.AwayFromZero);
        if (milliseconds >= (long)MillisecondsPerDay)
        {
            days += 1;
            milliseconds = 0;
        }

        var time = TimeSpan.FromMilliseconds(milliseconds);
        if (days == 0)
        {
            value = TimeOnlyDay + time;
            return true;
        }

        DateTime day;
        if (system == DateSystem.Date1904)
        {
            day = Base1904;
        }
        else if (days < 60)
        {
            day = Base1900;
        }
        else if (days == 60)
        {
            return false;
        }
        else
        {
            day = Base1900AfterLeapBug;
        }

        if (days > (DateTime.MaxValue.Date - day).TotalDays)
        {
            return false;
        }

        value = day.AddDays(days) + time;
        return true;
    }

    /// <summary>
    /// Converts a date and time to a serial. A value on <see cref="TimeOnlyDay"/> becomes its fraction of a day. False for a date before
    /// the date system's first day (1 January 1900, or 1 January 1904).
    /// </summary>
    public static bool TryFromDateTime(DateTime value, DateSystem system, out double serial)
    {
        serial = 0;
        double fraction = value.TimeOfDay.Ticks / (double)TimeSpan.TicksPerDay;
        if (IsTimeOnly(value))
        {
            serial = fraction;
            return true;
        }

        DateTime day = value.Date;
        double days;
        if (system == DateSystem.Date1904)
        {
            if (day < Base1904)
            {
                return false;
            }

            days = (day - Base1904).TotalDays;
        }
        else
        {
            if (day <= Base1900)
            {
                return false;
            }

            days = day < FirstDayAfterLeapBug ? (day - Base1900).TotalDays : (day - Base1900AfterLeapBug).TotalDays;
        }

        serial = days + fraction;
        return true;
    }
}
