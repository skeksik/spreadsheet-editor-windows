using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>The number format of each cell style (<c>cellXfs</c>) in a workbook's style sheet: its <c>numFmtId</c> and, for a custom id, its code.</summary>
internal sealed class NumberFormatTable
{
    private readonly uint[] formatIds;
    private readonly Dictionary<uint, string> customCodes;

    private NumberFormatTable(uint[] formatIds, Dictionary<uint, string> customCodes)
    {
        this.formatIds = formatIds;
        this.customCodes = customCodes;
    }

    public static NumberFormatTable Empty { get; } = new([], []);

    public static NumberFormatTable Read(X.Stylesheet? stylesheet)
    {
        if (stylesheet is null)
        {
            return Empty;
        }
        var codes = new Dictionary<uint, string>();
        foreach (var format in stylesheet.NumberingFormats?.Elements<X.NumberingFormat>() ?? [])
        {
            if (format.NumberFormatId?.Value is { } id && format.FormatCode?.Value is { } code)
            {
                codes[id] = code;
            }
        }
        var ids = new List<uint>();
        foreach (var format in stylesheet.CellFormats?.Elements<X.CellFormat>() ?? [])
        {
            ids.Add(format.NumberFormatId?.Value ?? 0);
        }
        return new NumberFormatTable([.. ids], codes);
    }

    /// <summary>The <c>numFmtId</c> of a cell style; 0 (General) for a style the sheet does not define.</summary>
    public int FormatId(uint styleIndex) => styleIndex < formatIds.Length ? (int)formatIds[styleIndex] : 0;

    /// <summary>
    /// The format code the workbook's <c>numFmts</c> gives for an id, or null when it gives none (a built-in format whose code is implied:
    /// ISO/IEC 29500-1 §18.8.30, as quoted in the Open XML SDK's NumberingFormat reference, read 2026-11-03).
    /// </summary>
    public string? FormatCode(int formatId) => formatId >= 0 && customCodes.TryGetValue((uint)formatId, out var code) ? code : null;
}
