namespace SpreadsheetEditor.Core.Values;

/// <summary>The date system of a workbook, which fixes the day a date serial counts from (<see cref="DateSerial"/>).</summary>
public enum DateSystem
{
    /// <summary>Serial 1 is 1 January 1900, and 1900 is counted as a leap year.</summary>
    Date1900,

    /// <summary>Serial 0 is 1 January 1904.</summary>
    Date1904,
}
