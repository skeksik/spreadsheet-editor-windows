namespace SpreadsheetEditor.Core.Model;

/// <summary>A table column: an entry on every card.</summary>
public sealed class Column
{
    public Column(string name) => Name = name;

    /// <summary>The header text; empty for a column without one.</summary>
    public string Name { get; set; }

    /// <summary>True when the column holds pictures (any of its cells had one when read, or the user added one).</summary>
    public bool IsImageColumn { get; set; }
}
