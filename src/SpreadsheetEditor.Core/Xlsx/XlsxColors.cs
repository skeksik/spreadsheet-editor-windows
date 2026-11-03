using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Resolves the colours of a workbook's style sheet: explicit ARGB, theme with tint, indexed, automatic (docs/spec.md §4.4, §9.5, §9.6).</summary>
internal sealed class XlsxColors
{
    private readonly IReadOnlyList<RgbColor> theme;
    private readonly IndexedPalette indexed;

    public XlsxColors(IReadOnlyList<RgbColor> theme, IndexedPalette indexed)
    {
        this.theme = theme;
        this.indexed = indexed;
    }

    /// <summary>The default theme palette and the standard indexed palette.</summary>
    public static XlsxColors Default { get; } = new(ThemePalette.Default, IndexedPalette.Standard);

    public static XlsxColors Read(X.Stylesheet? stylesheet, ThemePart? themePart) =>
        new(ThemePalette.Read(themePart), IndexedPalette.Read(stylesheet));

    /// <summary>
    /// A colour element (<c>color</c>, <c>fgColor</c>, ...): <c>auto</c> first, then <c>rgb</c>, <c>theme</c> (with <c>tint</c>) and
    /// <c>indexed</c>. A missing element or one without any of them is unresolved.
    /// </summary>
    public ResolvedColor Resolve(X.ColorType? color)
    {
        if (color is null)
        {
            return ResolvedColor.Unresolved;
        }
        if (color.Auto?.Value == true)
        {
            return ResolvedColor.Automatic;
        }
        if (color.Rgb?.Value is { } argb)
        {
            return Argb(argb);
        }
        if (color.Theme?.Value is { } index)
        {
            return index < theme.Count
                ? ResolvedColor.Of(ColorTint.Apply(theme[(int)index], color.Tint?.Value ?? 0))
                : ResolvedColor.Unresolved;
        }
        if (color.Indexed?.Value is { } legacy)
        {
            return indexed.Resolve(legacy);
        }
        return ResolvedColor.Unresolved;
    }

    /// <summary>
    /// An explicit colour: 8 hex digits, where <c>00000000</c> is no colour and otherwise the alpha is ignored, or 6 hex digits (docs/spec.md
    /// §4.4).
    /// </summary>
    public static ResolvedColor Argb(string text)
    {
        var hex = text.Trim();
        if (hex.Length is not (6 or 8) || !hex.All(char.IsAsciiHexDigit) || !RgbColor.TryParse(hex, out var color))
        {
            return ResolvedColor.Unresolved;
        }
        return hex == "00000000" ? ResolvedColor.None : ResolvedColor.Of(color);
    }
}
