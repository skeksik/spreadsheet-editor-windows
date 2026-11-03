namespace SpreadsheetEditor.Core.Model;

/// <summary>One worksheet's table: where it sits, its columns, its cards and the picture layout (docs/spec.md).</summary>
public sealed class Sheet
{
    public Sheet(string name) => Name = name;

    /// <summary>The worksheet's name, unique in the workbook.</summary>
    public string Name { get; set; }

    /// <summary>The single text above the header row, or null when the table has none.</summary>
    public string? TableName { get; set; }

    /// <summary>The 1-based worksheet column of <see cref="TableName"/>.</summary>
    public int TableNameColumn { get; set; } = 1;

    /// <summary>The 1-based worksheet column of the table's first column.</summary>
    public int FirstColumn { get; set; } = 1;

    /// <summary>The 1-based worksheet row of the header, or null when the table has no header row (its headers are hidden).</summary>
    public int? HeaderRow { get; set; } = 1;

    /// <summary>The 1-based worksheet row of the first card.</summary>
    public int FirstDataRow { get; set; } = 2;

    public List<Column> Columns { get; } = [];

    public List<Card> Cards { get; } = [];

    public ImageLayout ImageLayout { get; set; } = ImageLayout.Default;

    /// <summary>
    /// Width over height, captured when the user turned the aspect lock on (docs/spec.md §2.8), or null while it is off. Kept in memory
    /// only: never saved and not undoable (§8.1; whether to save it is open, §9.51). Set through <see cref="Editing.WorkbookEditor.SetAspectLock"/>.
    /// </summary>
    public double? LockedAspectRatio { get; internal set; }

    /// <summary>Where the table was in the worksheet it was read from, for the writer; null for a sheet not read from <c>.xlsx</c>.</summary>
    internal Xlsx.SheetSource? Source { get; set; }
}
