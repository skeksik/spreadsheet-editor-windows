using System.Globalization;

namespace SpreadsheetEditor.Core.Model;

/// <summary>An opaque sRGB colour, as a cell's fill or text colour.</summary>
public readonly record struct RgbColor(byte R, byte G, byte B)
{
    /// <summary>Parses <c>RRGGBB</c>, <c>#RRGGBB</c> or Excel's <c>AARRGGBB</c> (the alpha is dropped).</summary>
    public static bool TryParse(string? text, out RgbColor color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }
        var hex = text.Trim().TrimStart('#');
        if (hex.Length == 8)
        {
            hex = hex[2..];
        }
        if (hex.Length != 6 || !uint.TryParse(hex, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var value))
        {
            return false;
        }
        color = new RgbColor((byte)(value >> 16), (byte)(value >> 8), (byte)value);
        return true;
    }

    /// <summary><c>#RRGGBB</c>, upper case.</summary>
    public override string ToString() => $"#{R:X2}{G:X2}{B:X2}";
}
