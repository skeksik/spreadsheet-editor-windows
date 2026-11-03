namespace SpreadsheetEditor.Core.Model;

/// <summary>How a sheet lays out pictures in cells: one unit's size in pixels and how many units fit in a row (docs/spec.md).</summary>
public sealed record ImageLayout(int Width, int Height, int ImagesPerRow)
{
    public static ImageLayout Default { get; } = new(136, 180, 8);
}
