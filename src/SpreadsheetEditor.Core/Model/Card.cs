namespace SpreadsheetEditor.Core.Model;

/// <summary>One table row, or rows merged vertically into one, shown as a form: one entry per column.</summary>
public sealed class Card
{
    public List<Entry> Entries { get; } = [];

    /// <summary>How many worksheet rows the card took when read (1 for a new card).</summary>
    public int SourceRowSpan { get; set; } = 1;
}
