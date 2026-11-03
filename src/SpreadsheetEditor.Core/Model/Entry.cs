namespace SpreadsheetEditor.Core.Model;

/// <summary>A card's value in one column: its text, the typed value it was read as, its style and its pictures.</summary>
public sealed class Entry
{
    /// <summary>The text as shown and edited.</summary>
    public string Text { get; set; } = "";

    /// <summary>
    /// The cell's value as read, or null for an empty or new cell. While <see cref="Text"/> still equals its display text the writer keeps
    /// the typed value (a number stays a number); once the text changes, the writer stores the new text (docs/spec.md).
    /// </summary>
    public CellValue? Original { get; set; }

    /// <summary>The cell's style as read or set by the user; null keeps the worksheet's style.</summary>
    public CellStyle? Style { get; set; }

    public List<EntryImage> Images { get; } = [];
}
