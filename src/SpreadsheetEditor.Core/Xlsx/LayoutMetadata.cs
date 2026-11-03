using System.Globalization;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The picture layout a sheet stores in hidden cells of its last column (<c>XFD</c>), rows 1 to 3 (docs/spec.md §4.5), clamped to the
/// settings' ranges when read (§9.50).
/// </summary>
internal static class LayoutMetadata
{
    public const string HeightPrefix = "ImageHeight:";
    public const string ImagesPerRowPrefix = "ExcelImagesPerRow:";
    public const string WidthPrefix = "ImageWidth:";

    public const int MinWidth = 50;
    public const int MaxWidth = 1000;
    public const int MinHeight = 50;
    public const int MaxHeight = 500;

    /// <summary>The widest a cell's picture row may be, in pixels; pictures per row are at most this over the width (docs/spec.md §1.4).</summary>
    public const int MaxRowWidth = 1762;

    /// <summary>
    /// Reads the three cells in order, rows 1, 2, 3. The first that is missing or does not parse stops reading: it and the ones after it
    /// keep their defaults. Then every value is clamped to its range, pictures per row to the maximum for the (clamped) width.
    /// </summary>
    public static ImageLayout Read(string? row1, string? row2, string? row3)
    {
        var defaults = ImageLayout.Default;
        var height = defaults.Height;
        var perRow = defaults.ImagesPerRow;
        var width = defaults.Width;
        if (TryParse(row1, HeightPrefix, out var value))
        {
            height = value;
            if (TryParse(row2, ImagesPerRowPrefix, out value))
            {
                perRow = value;
                if (TryParse(row3, WidthPrefix, out value))
                {
                    width = value;
                }
            }
        }
        width = Math.Clamp(width, MinWidth, MaxWidth);
        height = Math.Clamp(height, MinHeight, MaxHeight);
        perRow = Math.Clamp(perRow, 1, MaxImagesPerRow(width));
        return new ImageLayout(width, height, perRow);
    }

    /// <summary>max(1, floor(1762 / width)) (docs/spec.md §1.4).</summary>
    public static int MaxImagesPerRow(int width) => Math.Max(1, MaxRowWidth / Math.Max(1, width));

    /// <summary>
    /// True when the text is one of the three metadata strings: a prefix, then an integer. Such cells are ignored when finding the table
    /// (docs/spec.md §9.1), in whichever column they are.
    /// </summary>
    public static bool IsMetadata(string? text) =>
        TryParse(text, HeightPrefix, out _) || TryParse(text, ImagesPerRowPrefix, out _) || TryParse(text, WidthPrefix, out _);

    /// <summary>The integer after the first <c>:</c> of a text that begins with the exact prefix.</summary>
    private static bool TryParse(string? text, string prefix, out int value)
    {
        value = 0;
        return text is not null
            && text.StartsWith(prefix, StringComparison.Ordinal)
            && int.TryParse(text.AsSpan(prefix.Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }
}
