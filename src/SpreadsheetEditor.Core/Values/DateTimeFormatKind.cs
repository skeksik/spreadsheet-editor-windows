namespace SpreadsheetEditor.Core.Values;

/// <summary>What a number format shows a number as (<see cref="NumberFormats.Classify"/>).</summary>
public enum DateTimeFormatKind
{
    /// <summary>Not a date or time format: the number is shown as a number.</summary>
    None,

    /// <summary>A date without a time.</summary>
    Date,

    /// <summary>A time of day or an elapsed time, without a date.</summary>
    Time,

    /// <summary>A date and a time.</summary>
    DateAndTime,
}
