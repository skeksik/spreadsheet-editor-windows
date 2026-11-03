using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Gives a cell a style with a changed number format, font, fill or alignment (docs/spec.md §5.5) without disturbing any other cell: the
/// cell's current <c>cellXfs</c> entry is copied, changed, and an equal existing entry (or font, fill, number format) is reused, else a new
/// one is appended. Entries are never changed or removed. The style sheet is loaded only when a style changes.
/// </summary>
internal sealed class StyleWriter
{
    /// <summary>The order of a font's children in the schema (ISO/IEC 29500-1 CT_Font, as written by the SDK's <c>Font</c> class).</summary>
    private static readonly string[] FontOrder =
        ["b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz", "color", "name", "family", "charset", "scheme"];

    private readonly WorkbookPart workbookPart;
    private readonly StyleTable original;
    private readonly int originalCount;
    private X.Stylesheet? stylesheet;

    /// <param name="workbookPart">The workbook part of the package being written.</param>
    /// <param name="original">The styles as read, to compare an entry's style with the style its cell has.</param>
    /// <param name="originalCount">The number of cell styles read; styles added by the writer are not in <paramref name="original"/>.</param>
    public StyleWriter(WorkbookPart workbookPart, StyleTable original, int originalCount)
    {
        this.workbookPart = workbookPart;
        this.original = original;
        this.originalCount = originalCount;
    }

    /// <summary>
    /// The style index for a cell whose style is <paramref name="current"/>, after applying a number format (null: unchanged), a style
    /// (null: unchanged) and, for text with a line break, wrap and top alignment. Returns <paramref name="current"/> when nothing changes.
    /// </summary>
    public uint Apply(uint current, NumberFormat? format, CellStyle? style, bool multiline)
    {
        style = style is null ? null : Changes(style, current);
        if (format is null && style is null && !multiline)
        {
            return current;
        }

        var sheet = Load();
        var formats = sheet.CellFormats!;
        var xfs = formats.Elements<X.CellFormat>().ToList();
        var original = current < xfs.Count ? xfs[(int)current] : xfs[0];
        var xf = (X.CellFormat)original.CloneNode(true);

        if (format is not null)
        {
            var id = NumberFormatId(sheet, format);
            if ((xf.NumberFormatId?.Value ?? 0) != id)
            {
                xf.NumberFormatId = id;
                xf.ApplyNumberFormat = true;
            }
        }

        if (style is not null)
        {
            ApplyFont(sheet, xf, style);
            ApplyFill(sheet, xf, style);
        }

        ApplyAlignment(xf, style, multiline);

        if (xf.OuterXml == original.OuterXml)
        {
            return current;
        }

        var index = xfs.FindIndex(x => x.OuterXml == xf.OuterXml);
        if (index >= 0)
        {
            return (uint)index;
        }

        formats.AppendChild(xf);
        formats.Count = (uint)(xfs.Count + 1);
        return (uint)xfs.Count;
    }

    /// <summary>
    /// The members of <paramref name="wanted"/> that differ from the style <paramref name="current"/> has as read, or null when none
    /// does: an entry's style as read (every read entry has one) changes nothing, and an edited member changes only itself. A member that
    /// leaves the value as it is (null, or no fill and no colour state) is never a change. A style the writer added is not compared: every
    /// member that is set applies.
    /// </summary>
    private CellStyle? Changes(CellStyle wanted, uint current)
    {
        if (current >= originalCount)
        {
            return wanted;
        }

        var have = original.Style(current);
        if (wanted == have)
        {
            return null;
        }

        var fill = (wanted.Fill is not null || wanted.NoFill) && (wanted.Fill != have.Fill || wanted.NoFill != have.NoFill);
        var color = (wanted.TextColor is not null || wanted.AutomaticTextColor)
            && (wanted.TextColor != have.TextColor || wanted.AutomaticTextColor != have.AutomaticTextColor);
        FontStyle? font = null;
        if (wanted.Font is { } f)
        {
            var h = have.Font;
            font = new FontStyle
            {
                Name = f.Name is not null && f.Name != h?.Name ? f.Name : null,
                Size = f.Size is not null && f.Size != h?.Size ? f.Size : null,
                Bold = f.Bold is not null && f.Bold != h?.Bold ? f.Bold : null,
                Italic = f.Italic is not null && f.Italic != h?.Italic ? f.Italic : null,
                Underline = f.Underline is not null && f.Underline != h?.Underline ? f.Underline : null,
                Strikethrough = f.Strikethrough is not null && f.Strikethrough != h?.Strikethrough ? f.Strikethrough : null,
            };
            if (font == new FontStyle())
            {
                font = null;
            }
        }

        var changes = new CellStyle
        {
            Fill = fill ? wanted.Fill : null,
            NoFill = fill && wanted.NoFill,
            TextColor = color ? wanted.TextColor : null,
            AutomaticTextColor = color && wanted.AutomaticTextColor,
            HorizontalAlignment = wanted.HorizontalAlignment is not null && wanted.HorizontalAlignment != have.HorizontalAlignment ? wanted.HorizontalAlignment : null,
            VerticalAlignment = wanted.VerticalAlignment is not null && wanted.VerticalAlignment != have.VerticalAlignment ? wanted.VerticalAlignment : null,
            WrapText = wanted.WrapText is not null && wanted.WrapText != have.WrapText ? wanted.WrapText : null,
            Font = font,
        };
        return changes == new CellStyle() ? null : changes;
    }

    private X.Stylesheet Load()
    {
        if (stylesheet is not null)
        {
            return stylesheet;
        }

        var part = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>(RelationshipIds.Next(workbookPart));
        stylesheet = part.Stylesheet ??= NewPackage.DefaultStylesheet();
        // A style sheet may lack any of the lists the writer needs; each gets the default first entry, as a new workbook has.
        var defaults = NewPackage.DefaultStylesheet();
        Ensure(stylesheet, defaults.Fonts!);
        Ensure(stylesheet, defaults.Fills!);
        Ensure(stylesheet, defaults.Borders!);
        Ensure(stylesheet, defaults.CellFormats!);
        return stylesheet;
    }

    private static void Ensure<T>(X.Stylesheet sheet, T list)
        where T : OpenXmlCompositeElement
    {
        var existing = sheet.GetFirstChild<T>();
        if (existing is null)
        {
            sheet.AddChild(list.CloneNode(true), throwOnError: false);
        }
        else if (!existing.HasChildren)
        {
            foreach (var child in list.ChildElements)
            {
                existing.AppendChild(child.CloneNode(true));
            }
        }
    }

    private static uint NumberFormatId(X.Stylesheet sheet, NumberFormat format)
    {
        if (format.Code is null)
        {
            return (uint)format.Id;
        }

        var list = sheet.NumberingFormats;
        var entries = list?.Elements<X.NumberingFormat>().ToList() ?? [];
        if (entries.Find(f => f.NumberFormatId?.Value == (uint)format.Id && f.FormatCode?.Value == format.Code) is { } same)
        {
            return same.NumberFormatId!.Value;
        }

        if (entries.Find(f => f.FormatCode?.Value == format.Code) is { } equal)
        {
            return equal.NumberFormatId!.Value;
        }

        // Custom formats take ids from 164 up (ECMA-376 §18.8.30, as quoted for NumberFormatTable).
        var id = Math.Max(164u, entries.Count == 0 ? 0 : entries.Max(f => f.NumberFormatId?.Value ?? 0) + 1);
        if (list is null)
        {
            list = new X.NumberingFormats();
            sheet.PrependChild(list);
        }

        list.AppendChild(new X.NumberingFormat { NumberFormatId = id, FormatCode = format.Code });
        list.Count = (uint)(entries.Count + 1);
        return id;
    }

    private static void ApplyFont(X.Stylesheet sheet, X.CellFormat xf, CellStyle style)
    {
        var font = style.Font;
        if (font is null && style.TextColor is null && !style.AutomaticTextColor)
        {
            return;
        }

        var fonts = sheet.Fonts!;
        var list = fonts.Elements<X.Font>().ToList();
        var id = xf.FontId?.Value ?? 0;
        var original = id < list.Count ? list[(int)id] : list[0];
        var children = original.ChildElements.Select(c => c.CloneNode(true)).ToList();

        if (font is not null)
        {
            SetFlag(children, "b", font.Bold, () => new X.Bold());
            SetFlag(children, "i", font.Italic, () => new X.Italic());
            SetFlag(children, "strike", font.Strikethrough, () => new X.Strike());
            if (font.Underline is { } underline)
            {
                children.RemoveAll(c => c.LocalName == "u");
                if (!string.Equals(underline, "none", StringComparison.OrdinalIgnoreCase))
                {
                    var element = new X.Underline();
                    if (!string.Equals(underline, "single", StringComparison.OrdinalIgnoreCase))
                    {
                        element.SetAttribute(new OpenXmlAttribute("val", "", underline));
                    }

                    children.Add(element);
                }
            }

            if (font.Size is { } size and > 0)
            {
                children.RemoveAll(c => c.LocalName == "sz");
                children.Add(new X.FontSize { Val = size });
            }

            if (!string.IsNullOrEmpty(font.Name))
            {
                children.RemoveAll(c => c.LocalName == "name");
                children.Add(new X.FontName { Val = font.Name });
            }
        }

        if (style.TextColor is { } color)
        {
            children.RemoveAll(c => c.LocalName == "color");
            children.Add(new X.Color { Rgb = Argb(color) });
        }
        else if (style.AutomaticTextColor)
        {
            // Automatic is a font without a colour, never an explicit black (§9.6).
            children.RemoveAll(c => c.LocalName == "color");
        }

        var changed = new X.Font();
        foreach (var child in children.OrderBy(c => Array.IndexOf(FontOrder, c.LocalName) is var i and >= 0 ? i : FontOrder.Length))
        {
            changed.AppendChild(child);
        }

        if (changed.OuterXml == original.OuterXml)
        {
            return;
        }

        xf.FontId = FindOrAdd(fonts, list, changed);
        xf.ApplyFont = true;
    }

    private static void SetFlag(List<OpenXmlElement> children, string name, bool? value, Func<OpenXmlElement> create)
    {
        if (value is not { } flag)
        {
            return;
        }

        children.RemoveAll(c => c.LocalName == name);
        if (flag)
        {
            children.Add(create());
        }
    }

    private static void ApplyFill(X.Stylesheet sheet, X.CellFormat xf, CellStyle style)
    {
        X.Fill changed;
        if (style.Fill is { } color)
        {
            changed = new X.Fill(new X.PatternFill(
                new X.ForegroundColor { Rgb = Argb(color) },
                new X.BackgroundColor { Indexed = 64 })
            { PatternType = X.PatternValues.Solid });
        }
        else if (style.NoFill)
        {
            changed = new X.Fill(new X.PatternFill { PatternType = X.PatternValues.None });
        }
        else
        {
            return;
        }

        var fills = sheet.Fills!;
        var list = fills.Elements<X.Fill>().ToList();
        var id = xf.FillId?.Value ?? 0;
        if (id < list.Count && list[(int)id].OuterXml == changed.OuterXml)
        {
            return;
        }

        xf.FillId = FindOrAdd(fills, list, changed);
        xf.ApplyFill = true;
    }

    private static void ApplyAlignment(X.CellFormat xf, CellStyle? style, bool multiline)
    {
        var horizontal = style?.HorizontalAlignment;
        var vertical = style?.VerticalAlignment;
        var wrap = style?.WrapText;
        if (horizontal is null && vertical is null && wrap is null && !multiline)
        {
            return;
        }

        var alignment = xf.Alignment is { } existing ? (X.Alignment)existing.CloneNode(true) : new X.Alignment();
        if (horizontal is not null)
        {
            alignment.SetAttribute(new OpenXmlAttribute("horizontal", "", horizontal));
        }

        if (vertical is not null)
        {
            alignment.SetAttribute(new OpenXmlAttribute("vertical", "", vertical));
        }

        if (wrap is { } flag)
        {
            alignment.WrapText = flag;
        }

        // Text with a line break shows its lines only when wrapped; it starts at the top unless a vertical alignment is set (§5.5).
        if (multiline)
        {
            alignment.WrapText = true;
            if (alignment.Vertical is null)
            {
                alignment.SetAttribute(new OpenXmlAttribute("vertical", "", "top"));
            }
        }

        if (xf.Alignment?.OuterXml == alignment.OuterXml)
        {
            return;
        }

        xf.Alignment = alignment;
        xf.ApplyAlignment = true;
    }

    private static uint FindOrAdd<T>(OpenXmlCompositeElement list, List<T> entries, T element)
        where T : OpenXmlElement
    {
        var index = entries.FindIndex(e => e.OuterXml == element.OuterXml);
        if (index >= 0)
        {
            return (uint)index;
        }

        list.AppendChild(element);
        list.SetAttribute(new OpenXmlAttribute("count", "", (entries.Count + 1).ToString(CultureInfo.InvariantCulture)));
        return (uint)entries.Count;
    }

    private static string Argb(RgbColor color) => $"FF{color.R:X2}{color.G:X2}{color.B:X2}";
}
