using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The <see cref="CellStyle"/> of each cell style (<c>cellXfs</c>) in a workbook's style sheet: fill, text colour, font and alignment
/// (docs/spec.md §4.2, §4.4, §9.5, §9.6). It holds nothing but the style sheet and the theme, so the writer can build one from the package
/// it saves into and compare an entry's style with the style its cell has there. Styles are built on first use and shared: a
/// <see cref="CellStyle"/> is immutable.
/// </summary>
internal sealed class StyleTable
{
    private readonly X.CellFormat[] formats;
    private readonly X.Font[] fonts;
    private readonly X.Fill[] fills;
    private readonly XlsxColors colors;
    private readonly CellStyle?[] cache;
    private CellStyle? fallback;

    private StyleTable(X.CellFormat[] formats, X.Font[] fonts, X.Fill[] fills, XlsxColors colors)
    {
        this.formats = formats;
        this.fonts = fonts;
        this.fills = fills;
        this.colors = colors;
        cache = new CellStyle?[formats.Length];
    }

    /// <summary>A table for a workbook without a style sheet: every index gives the style of a cell with no fill and an automatic font.</summary>
    public static StyleTable Empty { get; } = new([], [], [], XlsxColors.Default);

    /// <summary>Reads the style sheet's cell styles, fonts, fills and indexed colours, and the theme's colours.</summary>
    /// <param name="stylesheet">The workbook's style sheet, or null when it has none.</param>
    /// <param name="theme">The workbook's theme part, or null: theme colours then come from the default palette.</param>
    public static StyleTable Read(X.Stylesheet? stylesheet, ThemePart? theme)
    {
        if (stylesheet is null)
        {
            return new StyleTable([], [], [], XlsxColors.Read(null, theme));
        }
        return new StyleTable(
            [.. stylesheet.CellFormats?.Elements<X.CellFormat>() ?? []],
            [.. stylesheet.Fonts?.Elements<X.Font>() ?? []],
            [.. stylesheet.Fills?.Elements<X.Fill>() ?? []],
            XlsxColors.Read(stylesheet, theme));
    }

    /// <summary>The style of cell style <paramref name="index"/>; an index the style sheet does not define gives the style of index 0.</summary>
    public CellStyle Style(uint index)
    {
        if (index >= formats.Length)
        {
            return index == 0 || formats.Length == 0 ? fallback ??= Build(null) : Style(0);
        }
        return cache[index] ??= Build(formats[index]);
    }

    private CellStyle Build(X.CellFormat? format)
    {
        var font = Element(fonts, format?.FontId?.Value ?? 0);
        var fill = Fill(Element(fills, format?.FillId?.Value ?? 0));
        var text = font is null ? ResolvedColor.Automatic : TextColor(font.Color);
        var alignment = format?.Alignment;
        return new CellStyle
        {
            Fill = fill.Kind == ColorKind.Rgb ? fill.Rgb : null,
            NoFill = fill.Kind == ColorKind.None,
            TextColor = text.Kind == ColorKind.Rgb ? text.Rgb : null,
            AutomaticTextColor = text.Kind == ColorKind.Automatic,
            HorizontalAlignment = alignment?.Horizontal?.InnerText,
            VerticalAlignment = alignment?.Vertical?.InnerText,
            WrapText = alignment?.WrapText?.Value,
            Font = font is null ? null : Font(font),
        };
    }

    /// <summary>
    /// A pattern fill whose type is not <c>none</c> gives its foreground colour; no fill, an empty pattern fill or <c>patternType="none"</c>
    /// give none. A gradient fill, a missing foreground colour or an automatic one cannot be shown as one colour: unresolved.
    /// </summary>
    private ResolvedColor Fill(X.Fill? fill)
    {
        if (fill is null)
        {
            return ResolvedColor.None;
        }
        if (fill.PatternFill is not { } pattern)
        {
            return fill.GradientFill is null ? ResolvedColor.None : ResolvedColor.Unresolved;
        }
        if (pattern.PatternType?.Value is not { } type || type == X.PatternValues.None)
        {
            return ResolvedColor.None;
        }
        var color = colors.Resolve(pattern.ForegroundColor);
        return color.Kind == ColorKind.Automatic ? ResolvedColor.Unresolved : color;
    }

    /// <summary>A font without a colour has the automatic colour (§9.6); a fully transparent one cannot be shown: unresolved.</summary>
    private ResolvedColor TextColor(X.Color? color)
    {
        if (color is null)
        {
            return ResolvedColor.Automatic;
        }
        var resolved = colors.Resolve(color);
        return resolved.Kind == ColorKind.None ? ResolvedColor.Unresolved : resolved;
    }

    private static FontStyle Font(X.Font font) => new()
    {
        Name = font.FontName?.Val?.Value,
        Size = font.FontSize?.Val?.Value,
        Bold = Flag(font.Bold?.Val, font.Bold is not null),
        Italic = Flag(font.Italic?.Val, font.Italic is not null),
        Strikethrough = Flag(font.Strike?.Val, font.Strike is not null),
        Underline = Underline(font.Underline),
    };

    /// <summary>
    /// A boolean font property: absent is false; present without <c>val</c> (Excel writes a bare <c>&lt;b/&gt;</c>) is true, as the
    /// property's <c>val</c> defaults to true (not verified against the schema text).
    /// </summary>
    private static bool Flag(DocumentFormat.OpenXml.BooleanValue? value, bool present) => present && (value?.Value ?? true);

    /// <summary>A bare <c>&lt;u/&gt;</c> is a single underline; <c>val="none"</c> is no underline.</summary>
    private static string? Underline(X.Underline? underline)
    {
        if (underline is null)
        {
            return null;
        }
        var kind = underline.Val?.InnerText;
        return string.IsNullOrEmpty(kind) ? "single" : kind == "none" ? null : kind;
    }

    private static T? Element<T>(T[] elements, uint index)
        where T : class => index < elements.Length ? elements[index] : null;
}
