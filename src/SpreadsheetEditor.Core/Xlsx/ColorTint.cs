using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Applies a theme colour's tint to its lightness in HSL, exactly as docs/spec.md §4.4 step 4 gives the algorithm.</summary>
internal static class ColorTint
{
    public static RgbColor Apply(RgbColor color, double tint)
    {
        if (tint == 0 || double.IsNaN(tint))
        {
            return color;
        }
        // The tint is a value from -1 to 1 (docs/spec.md §4.4); a value outside that range in a damaged file is clamped.
        tint = Math.Clamp(tint, -1, 1);
        var r = color.R / 255.0;
        var g = color.G / 255.0;
        var b = color.B / 255.0;
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var l = (max + min) / 2;
        double h = 0;
        double s = 0;
        if (max != min)
        {
            var d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            h = max == r ? ((g - b) / d) + (g < b ? 6 : 0)
                : max == g ? ((b - r) / d) + 2
                : ((r - g) / d) + 4;
            h /= 6;
        }

        l = tint < 0 ? l * (1 + tint) : (l * (1 - tint)) + tint;

        if (s == 0)
        {
            r = g = b = l;
        }
        else
        {
            var q = l < 0.5 ? l * (1 + s) : l + s - (l * s);
            var p = (2 * l) - q;
            r = Hue(p, q, h + (1.0 / 3));
            g = Hue(p, q, h);
            b = Hue(p, q, h - (1.0 / 3));
        }
        return new RgbColor(Channel(r), Channel(g), Channel(b));
    }

    private static double Hue(double p, double q, double t)
    {
        if (t < 0)
        {
            t += 1;
        }
        if (t > 1)
        {
            t -= 1;
        }
        if (t < 1.0 / 6)
        {
            return p + ((q - p) * 6 * t);
        }
        if (t < 1.0 / 2)
        {
            return q;
        }
        if (t < 2.0 / 3)
        {
            return p + ((q - p) * ((2.0 / 3) - t) * 6);
        }
        return p;
    }

    private static byte Channel(double x) => (byte)Math.Clamp(Math.Floor((x * 255) + 0.5), 0, 255);
}
