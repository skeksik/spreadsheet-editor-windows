namespace SpreadsheetEditor.Core.Imaging;

/// <summary>Column widths in the worksheet's character-width units and in pixels (docs/spec.md §5.2).</summary>
public static class ColumnWidths
{
    /// <summary>The width of a column without an explicit width.</summary>
    public const double DefaultWidth = 8.43;

    /// <summary>The widest a picture column is made, in pixels; also the limit on pictures per row times picture width.</summary>
    public const int MaxPx = 1762;

    /// <summary>Pixels to a column width: px / 12 up to 12 px, (px − 5) / 7 above.</summary>
    public static double PxToWidth(double px) => px <= 12 ? px / 12 : (px - 5) / 7;

    /// <summary>A column width to pixels: width × 12 up to width 1, width × 7 + 5 above.</summary>
    public static double WidthToPx(double width) => width <= 1 ? width * 12 : (width * 7) + 5;

    /// <summary>
    /// The width a picture column gets for a target in pixels, or null when the column stays as it is. The target is capped at
    /// <see cref="MaxPx"/>, and a column is only ever widened.
    /// </summary>
    /// <param name="currentWidth">The column's explicit width, or null for <see cref="DefaultWidth"/>.</param>
    /// <param name="targetPx">The width the column's pictures need, in pixels.</param>
    public static double? Widen(double? currentWidth, int targetPx)
    {
        var target = Math.Min(targetPx, MaxPx);
        return target > WidthToPx(currentWidth ?? DefaultWidth) ? PxToWidth(target) : null;
    }
}
