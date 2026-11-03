using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Csv;

/// <summary>A workbook read from a .csv file, with the format the file used, for the caller to keep and save the sheet with.</summary>
public sealed record CsvReadResult(Workbook Workbook, CsvFormat Format);
