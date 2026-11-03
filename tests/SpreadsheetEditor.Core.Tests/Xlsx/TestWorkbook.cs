using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>Builds small <c>.xlsx</c> packages in code with the Open XML SDK, for reader tests. Every value in them is made up.</summary>
internal sealed class TestWorkbook
{
    private readonly List<(string Name, TestSheet? Sheet, bool Hidden)> sheets = [];
    private readonly List<X.SharedStringItem> sharedStrings = [];
    private readonly List<(uint Id, string Code)> customFormats = [];
    private readonly List<uint> cellFormats = [0];

    public bool Date1904 { get; set; }

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
        cellFormats.Add(formatId);
        return (uint)(cellFormats.Count - 1);
    }

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
            workbookPart.Workbook = root;

            var stylesheet = new X.Stylesheet();
            if (customFormats.Count > 0)
            {
                stylesheet.AppendChild(new X.NumberingFormats(
                    customFormats.Select(f => new X.NumberingFormat { NumberFormatId = f.Id, FormatCode = f.Code })));
            }
            stylesheet.AppendChild(new X.Fonts(new X.Font()));
            stylesheet.AppendChild(new X.Fills(new X.Fill()));
            stylesheet.AppendChild(new X.Borders(new X.Border()));
            stylesheet.AppendChild(new X.CellFormats(cellFormats.Select(id => new X.CellFormat { NumberFormatId = id })));
            workbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;

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
