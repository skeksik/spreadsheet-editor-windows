using SpreadsheetEditor.Core.Model;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The legacy indexed colours a style's <c>indexed</c> attribute refers to (docs/spec.md §9.5). The standard palette is the default
/// mapping of ECMA-376 Part 1 / ISO/IEC 29500-1 §18.8.27 (indexedColors), as quoted in the Open XML SDK's <c>IndexedColors</c> class
/// reference (learn.microsoft.com, read 2026-11-03): "Here is the table of default mappings from indexed color value to ARGB value. Note
/// that 0-7 are redundant of 8-15 to preserve backwards compatibility." It ends with "indexed="64" System Foreground" and "indexed="65"
/// System Background", which have no ARGB value and are read as the automatic colour. The same text says "When the color palette has been
/// modified from default, then the entire color palette is written out": the style sheet's <c>indexedColors</c> then replaces the
/// palette.
/// </summary>
internal sealed class IndexedPalette
{
    /// <summary>The system foreground colour (automatic).</summary>
    public const int SystemForeground = 64;

    /// <summary>The system background colour (automatic).</summary>
    public const int SystemBackground = 65;

    /// <summary>The 64 standard colours (RRGGBB; the table's ARGB values all have alpha 00, which a palette entry ignores).</summary>
    private static readonly uint[] StandardValues =
    [
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, // 0-7
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, // 8-15
        0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080, // 16-23
        0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, // 24-31
        0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF, // 32-39
        0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99, // 40-47
        0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696, // 48-55
        0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333, // 56-63
    ];

    private readonly RgbColor?[] colors;

    private IndexedPalette(RgbColor?[] colors) => this.colors = colors;

    public static IndexedPalette Standard { get; } = new([.. StandardValues.Select(v => (RgbColor?)new RgbColor((byte)(v >> 16), (byte)(v >> 8), (byte)v))]);

    /// <summary>The style sheet's <c>indexedColors</c> when it has one, else the standard palette.</summary>
    public static IndexedPalette Read(X.Stylesheet? stylesheet)
    {
        var custom = stylesheet?.Colors?.IndexedColors?.Elements<X.RgbColor>().ToList();
        if (custom is null || custom.Count == 0)
        {
            return Standard;
        }
        // An entry that cannot be parsed resolves to nothing; indexes the custom palette does not reach fall back to the standard one.
        var colors = new RgbColor?[Math.Max(custom.Count, StandardValues.Length)];
        for (var i = 0; i < colors.Length; i++)
        {
            colors[i] = i < custom.Count
                ? (RgbColor.TryParse(custom[i].Rgb?.Value, out var color) ? color : null)
                : Standard.colors[i];
        }
        return new IndexedPalette(colors);
    }

    public ResolvedColor Resolve(uint index) => index switch
    {
        SystemForeground or SystemBackground => ResolvedColor.Automatic,
        _ when index < colors.Length && colors[index] is { } color => ResolvedColor.Of(color),
        _ => ResolvedColor.Unresolved,
    };
}
