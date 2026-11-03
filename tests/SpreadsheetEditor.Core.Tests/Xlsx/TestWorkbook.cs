using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>Builds small <c>.xlsx</c> packages in code with the Open XML SDK, for reader tests. Every value in them is made up.</summary>
internal sealed class TestWorkbook
{
    private readonly List<(string Name, TestSheet? Sheet, bool Hidden)> sheets = [];
    private readonly List<X.SharedStringItem> sharedStrings = [];
    private readonly List<(uint Id, string Code)> customFormats = [];
    private readonly List<X.CellFormat> cellFormats = [new X.CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0 }];
    private readonly List<X.Font> fonts = [new X.Font(new X.FontSize { Val = 11 }, new X.FontName { Val = "Calibri" })];

    // Fills 0 and 1 are the two every workbook Excel writes has: none and gray125.
    private readonly List<X.Fill> fills =
    [
        new X.Fill(new X.PatternFill { PatternType = X.PatternValues.None }),
        new X.Fill(new X.PatternFill { PatternType = X.PatternValues.Gray125 }),
    ];

    private readonly List<(string Name, string Text, uint? LocalSheetId)> definedNames = [];

    public bool Date1904 { get; set; }

    /// <summary>When set, the package has a calculation chain listing cell A1 of the first sheet.</summary>
    public bool CalculationChain { get; set; }

    /// <summary>Adds a defined name; <paramref name="localSheetId"/> makes it local to the sheet at that index.</summary>
    public void AddDefinedName(string name, string text, uint? localSheetId = null) => definedNames.Add((name, text, localSheetId));

    /// <summary>The style sheet's custom indexed palette, or null for the standard one.</summary>
    public X.IndexedColors? IndexedColors { get; set; }

    /// <summary>The theme's colour scheme; null writes no theme part.</summary>
    public A.ColorScheme? ThemeColors { get; set; }

    /// <summary>Replaces the font of the default cell style (index 0).</summary>
    public X.Font DefaultFont
    {
        get => fonts[0];
        set => fonts[0] = value;
    }

    /// <summary>Adds a worksheet; a hidden one has <c>state="hidden"</c> in the workbook's sheet list.</summary>
    public TestSheet AddSheet(string name, bool hidden = false)
    {
        var sheet = new TestSheet(this);
        sheets.Add((name, sheet, hidden));
        return sheet;
    }

    /// <summary>Adds a chart sheet (no grid) with this name.</summary>
    public void AddChartSheet(string name) => sheets.Add((name, null, false));

    /// <summary>Adds a cell style with this number format and returns its index; <paramref name="code"/> defines a custom format.</summary>
    public uint AddNumberFormat(uint formatId, string? code = null)
    {
        if (code is not null)
        {
            customFormats.Add((formatId, code));
        }
        cellFormats.Add(new X.CellFormat { NumberFormatId = formatId, FontId = 0, FillId = 0 });
        return (uint)(cellFormats.Count - 1);
    }

    /// <summary>Adds a cell style with this font, fill and alignment (each optional) and returns its index.</summary>
    public uint AddStyle(X.Font? font = null, X.Fill? fill = null, X.Alignment? alignment = null)
    {
        var format = new X.CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0 };
        if (font is not null)
        {
            fonts.Add(font);
            format.FontId = (uint)(fonts.Count - 1);
        }
        if (fill is not null)
        {
            fills.Add(fill);
            format.FillId = (uint)(fills.Count - 1);
        }
        if (alignment is not null)
        {
            format.Alignment = alignment;
        }
        cellFormats.Add(format);
        return (uint)(cellFormats.Count - 1);
    }

    /// <summary>A solid fill of this colour element (<c>fgColor</c>).</summary>
    public static X.Fill Solid(X.ForegroundColor color) =>
        new(new X.PatternFill(color) { PatternType = X.PatternValues.Solid });

    internal int AddSharedString(X.SharedStringItem item)
    {
        sharedStrings.Add(item);
        return sharedStrings.Count - 1;
    }

    public byte[] ToBytes()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            var sheetList = new X.Sheets();
            var root = new X.Workbook();
            if (Date1904)
            {
                root.AppendChild(new X.WorkbookProperties { Date1904 = true });
            }
            root.AppendChild(sheetList);
            if (definedNames.Count > 0)
            {
                root.AppendChild(new X.DefinedNames(definedNames.Select(n => new X.DefinedName(n.Text) { Name = n.Name, LocalSheetId = n.LocalSheetId })));
            }
            workbookPart.Workbook = root;
            if (CalculationChain)
            {
                workbookPart.AddNewPart<CalculationChainPart>().CalculationChain =
                    new X.CalculationChain(new X.CalculationCell { CellReference = "A1", SheetId = 1 });
            }

            var stylesheet = new X.Stylesheet();
            if (customFormats.Count > 0)
            {
                stylesheet.AppendChild(new X.NumberingFormats(
                    customFormats.Select(f => new X.NumberingFormat { NumberFormatId = f.Id, FormatCode = f.Code })));
            }
            stylesheet.AppendChild(new X.Fonts(fonts.Select(f => (X.Font)f.CloneNode(true))));
            stylesheet.AppendChild(new X.Fills(fills.Select(f => (X.Fill)f.CloneNode(true))));
            stylesheet.AppendChild(new X.Borders(new X.Border()));
            stylesheet.AppendChild(new X.CellFormats(cellFormats.Select(f => (X.CellFormat)f.CloneNode(true))));
            if (IndexedColors is not null)
            {
                stylesheet.AppendChild(new X.Colors((X.IndexedColors)IndexedColors.CloneNode(true)));
            }
            workbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;

            if (ThemeColors is not null)
            {
                workbookPart.AddNewPart<ThemePart>().Theme = new A.Theme(
                    new A.ThemeElements((A.ColorScheme)ThemeColors.CloneNode(true), new A.FontScheme(), new A.FormatScheme()))
                { Name = "Test" };
            }

            if (sharedStrings.Count > 0)
            {
                workbookPart.AddNewPart<SharedStringTablePart>().SharedStringTable =
                    new X.SharedStringTable(sharedStrings.Select(s => (X.SharedStringItem)s.CloneNode(true)));
            }

            uint sheetId = 1;
            foreach (var (name, sheet, hidden) in sheets)
            {
                OpenXmlPart part;
                if (sheet is null)
                {
                    var chartPart = workbookPart.AddNewPart<ChartsheetPart>();
                    chartPart.Chartsheet = new X.Chartsheet();
                    part = chartPart;
                }
                else
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = sheet.ToWorksheet();
                    sheet.AddDrawing(worksheetPart);
                    part = worksheetPart;
                }
                var entry = new X.Sheet { Id = workbookPart.GetIdOfPart(part), SheetId = sheetId++, Name = name };
                if (hidden)
                {
                    entry.State = X.SheetStateValues.Hidden;
                }
                sheetList.AppendChild(entry);
            }
        }
        return stream.ToArray();
    }

    public MemoryStream ToStream() => new(ToBytes());
}
