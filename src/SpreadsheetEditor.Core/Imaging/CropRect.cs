namespace SpreadsheetEditor.Core.Imaging;

/// <summary>
/// A source rectangle: how much to cut from each edge of a picture, in units of 1/100000 of its width (left, right) or height (top,
/// bottom). All zero means no source rectangle.
/// </summary>
public readonly record struct CropRect(int Left, int Top, int Right, int Bottom)
{
    public static CropRect None => default;

    public bool IsNone => this == default;
}
