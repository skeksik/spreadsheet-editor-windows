namespace SpreadsheetEditor.Core.Model;

/// <summary>A table column: an entry on every card.</summary>
public sealed class Column
{
    public Column(string name) => Name = name;

    /// <summary>The header text; empty for a column without one.</summary>
    public string Name { get; set; }

    /// <summary>True when the column holds pictures (any of its cells had one when read, or the user added one).</summary>
    public bool IsImageColumn { get; set; }

    /// <summary>
    /// The column's 0-based position in the table when it was read from <c>.xlsx</c>, or null for a column added since. The writer uses it to
    /// find the worksheet column an entry's values and header came from after entries were removed or reordered.
    /// </summary>
    internal int? SourceIndex { get; set; }
}
