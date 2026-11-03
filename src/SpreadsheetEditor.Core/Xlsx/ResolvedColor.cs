using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>A colour from the workbook's XML, resolved: its kind and, for <see cref="ColorKind.Rgb"/>, its value.</summary>
internal readonly record struct ResolvedColor(ColorKind Kind, RgbColor Rgb)
{
    public static ResolvedColor Unresolved => new(ColorKind.Unresolved, default);

    public static ResolvedColor None => new(ColorKind.None, default);

    public static ResolvedColor Automatic => new(ColorKind.Automatic, default);

    public static ResolvedColor Of(RgbColor color) => new(ColorKind.Rgb, color);
}
