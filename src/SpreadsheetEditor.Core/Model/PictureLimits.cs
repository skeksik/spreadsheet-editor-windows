namespace SpreadsheetEditor.Core.Model;

/// <summary>The ranges of a sheet's picture layout settings (docs/spec.md §2.8, §9.50), for the editor, the reader and the app's boxes.</summary>
public static class PictureLimits
{
    public const int MinWidth = 50;
    public const int MaxWidth = 1000;
    public const int MinHeight = 50;
    public const int MaxHeight = 500;

    /// <summary>The widest a cell's picture row may be, in pixels; pictures per row are at most this over the width (docs/spec.md §1.4).</summary>
    public const int MaxRowWidth = 1762;

    /// <summary>max(1, floor(1762 / width)) (docs/spec.md §1.4).</summary>
    public static int MaxImagesPerRow(int width) => Math.Max(1, MaxRowWidth / Math.Max(1, width));
}
