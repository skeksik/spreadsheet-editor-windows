namespace SpreadsheetEditor.Core.Imaging;

/// <summary>The centre crop of docs/spec.md §5.4 step 3.</summary>
public static class CenterCrop
{
    /// <summary>The units of a <see cref="CropRect"/>: 100000 is the whole width or height.</summary>
    public const int Scale = 100000;

    /// <summary>
    /// The source rectangle that centre-crops an image of <paramref name="imageWidth"/> x <paramref name="imageHeight"/> pixels to the
    /// aspect <paramref name="targetWidth"/> / <paramref name="targetHeight"/>, truncated to integers. Same aspect, or a size that is
    /// not positive, gives <see cref="CropRect.None"/>.
    /// </summary>
    public static CropRect Compute(int imageWidth, int imageHeight, int targetWidth, int targetHeight)
    {
        if (imageWidth <= 0 || imageHeight <= 0 || targetWidth <= 0 || targetHeight <= 0)
        {
            return CropRect.None;
        }

        // Compare imageWidth / imageHeight with targetWidth / targetHeight exactly.
        var imageSide = (long)imageWidth * targetHeight;
        var targetSide = (long)imageHeight * targetWidth;
        if (imageSide == targetSide)
        {
            return CropRect.None;
        }

        var target = (double)targetWidth / targetHeight;
        if (imageSide > targetSide)
        {
            // Too wide.
            var newWidth = imageHeight * target;
            var crop = imageWidth - newWidth;
            var left = crop / 2;
            var right = crop - left;
            return new CropRect(ToUnits(left, imageWidth), 0, ToUnits(right, imageWidth), 0);
        }

        // Too tall.
        var newHeight = imageWidth / target;
        var cropHeight = imageHeight - newHeight;
        var top = cropHeight / 2;
        var bottom = cropHeight - top;
        return new CropRect(0, ToUnits(top, imageHeight), 0, ToUnits(bottom, imageHeight));
    }

    private static int ToUnits(double pixels, int size) => (int)Math.Truncate(pixels / size * Scale);
}
