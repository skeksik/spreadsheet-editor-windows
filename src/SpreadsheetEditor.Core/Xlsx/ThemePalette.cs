using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>The twelve theme colours a style's <c>theme</c> index refers to, read from the workbook's theme (docs/spec.md §4.4 step 1).</summary>
internal static class ThemePalette
{
    /// <summary>The slots of the theme's colour scheme in index order: lt1, dk1, lt2, dk2, the six accents, hlink, folHlink.</summary>
    private static readonly string[] Slots = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

    private static readonly RgbColor White = new(0xFF, 0xFF, 0xFF);

    /// <summary>The palette used when the workbook has no theme, or its theme cannot be read (docs/spec.md §4.4 step 1).</summary>
    public static IReadOnlyList<RgbColor> Default { get; } =
    [
        new(0xFF, 0xFF, 0xFF), new(0x00, 0x00, 0x00), new(0xE7, 0xE6, 0xE6), new(0x44, 0x54, 0x6A),
        new(0x5B, 0x9B, 0xD5), new(0xED, 0x7D, 0x31), new(0xA5, 0xA5, 0xA5), new(0xFF, 0xC0, 0x00),
        new(0x44, 0x72, 0xC4), new(0x70, 0xAD, 0x47), new(0x05, 0x63, 0xC1), new(0x95, 0x4F, 0x72),
    ];

    public static int Count => Slots.Length;

    /// <summary>
    /// The theme's colours: each slot's <c>srgbClr</c> value, a system colour's <c>lastClr</c> (white without one), white for a missing or
    /// unreadable slot. The default palette when there is no theme or no colour scheme.
    /// </summary>
    public static IReadOnlyList<RgbColor> Read(ThemePart? part)
    {
        var scheme = part?.Theme?.ThemeElements?.ColorScheme;
        if (scheme is null)
        {
            return Default;
        }
        var colors = new RgbColor[Slots.Length];
        for (var i = 0; i < Slots.Length; i++)
        {
            var slot = scheme.ChildElements.FirstOrDefault(e => e.LocalName == Slots[i]);
            colors[i] = slot is null ? White : SlotColor(slot);
        }
        return colors;
    }

    private static RgbColor SlotColor(OpenXmlElement slot)
    {
        foreach (var child in slot.ChildElements)
        {
            var attribute = child.LocalName switch
            {
                "srgbClr" => "val",
                "sysClr" => "lastClr",
                _ => null,
            };
            if (attribute is null)
            {
                continue;
            }
            var value = child.GetAttributes().FirstOrDefault(a => a.LocalName == attribute && string.IsNullOrEmpty(a.NamespaceUri)).Value;
            return RgbColor.TryParse(value, out var color) && value!.Trim().Length == 6 ? color : White;
        }
        return White;
    }
}
