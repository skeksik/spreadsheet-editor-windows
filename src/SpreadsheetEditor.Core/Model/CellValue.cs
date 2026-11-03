namespace SpreadsheetEditor.Core.Model;

public enum CellValueKind
{
    Text,
    Number,
    Boolean,
    Date,
}

/// <summary>A typed cell value as read from a file, with the text the app showed for it.</summary>
public sealed record CellValue(CellValueKind Kind, string DisplayText, double? Number = null, bool? Boolean = null, DateTime? Date = null);
