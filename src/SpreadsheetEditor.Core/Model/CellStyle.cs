namespace SpreadsheetEditor.Core.Model;

/// <summary>The style an entry shows and writes back. A null member leaves the worksheet's value unchanged.</summary>
public sealed record CellStyle
{
    public RgbColor? Fill { get; init; }

    /// <summary>True when the user removed the fill; <see cref="Fill"/> is then null.</summary>
    public bool NoFill { get; init; }

    public RgbColor? TextColor { get; init; }

    /// <summary>
    /// True when the text colour is automatic: the font has no colour, or one marked automatic (docs/spec.md §9.6). <see cref="TextColor"/>
    /// is then null.
    /// </summary>
    public bool AutomaticTextColor { get; init; }

    /// <summary>Excel's horizontal alignment name (left, center, right, ...), or null.</summary>
    public string? HorizontalAlignment { get; init; }

    /// <summary>Excel's vertical alignment name (top, center, bottom, ...), or null.</summary>
    public string? VerticalAlignment { get; init; }

    public bool? WrapText { get; init; }

    public FontStyle? Font { get; init; }
}

/// <summary>A cell's font. A null member leaves the worksheet's value unchanged.</summary>
public sealed record FontStyle
{
    public string? Name { get; init; }

    public double? Size { get; init; }

    public bool? Bold { get; init; }

    public bool? Italic { get; init; }

    /// <summary>Excel's underline name (single, double, ...), or null.</summary>
    public string? Underline { get; init; }

    public bool? Strikethrough { get; init; }
}
