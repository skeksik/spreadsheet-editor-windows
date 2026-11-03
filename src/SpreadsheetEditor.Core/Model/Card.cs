namespace SpreadsheetEditor.Core.Model;

/// <summary>One table row, or rows merged vertically into one, shown as a form: one entry per column.</summary>
public sealed class Card
{
    public List<Entry> Entries { get; } = [];

    /// <summary>How many worksheet rows the card took when read (1 for a new card).</summary>
    public int SourceRowSpan { get; set; } = 1;

    /// <summary>The 1-based worksheet row the card started on when read, or null for a card that was not read from <c>.xlsx</c>.</summary>
    internal int? SourceFirstRow { get; set; }
}
