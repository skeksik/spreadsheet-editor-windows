using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>Looks inside a written <c>.xlsx</c> package, for writer tests.</summary>
internal sealed class XlsxPackage : IDisposable
{
    private readonly MemoryStream stream;

    public XlsxPackage(byte[] bytes)
    {
        stream = new MemoryStream(bytes, writable: false);
        Document = SpreadsheetDocument.Open(stream, false);
    }

    public SpreadsheetDocument Document { get; }

    public WorkbookPart WorkbookPart => Document.WorkbookPart!;

    public X.Workbook Workbook => WorkbookPart.Workbook!;

    public List<string> SheetNames => Workbook.Sheets!.Elements<X.Sheet>().Select(s => s.Name!.Value!).ToList();

    public X.Sheet Entry(string sheetName) => Workbook.Sheets!.Elements<X.Sheet>().Single(s => s.Name?.Value == sheetName);

    public WorksheetPart Part(string sheetName) => (WorksheetPart)WorkbookPart.GetPartById(Entry(sheetName).Id!.Value!);

    public X.Worksheet Worksheet(string sheetName) => Part(sheetName).Worksheet!;

    public X.Cell? Cell(string sheetName, string reference) =>
        Worksheet(sheetName).Descendants<X.Cell>().SingleOrDefault(c => c.CellReference?.Value == reference);

    /// <summary>The shared string a string cell holds, or the cell's raw value.</summary>
    public string? Text(string sheetName, string reference)
    {
        var cell = Cell(sheetName, reference);
        if (cell?.CellValue is null)
        {
            return null;
        }

        if (cell.DataType?.Value == X.CellValues.SharedString)
        {
            var items = WorkbookPart.SharedStringTablePart!.SharedStringTable!.Elements<X.SharedStringItem>().ToList();
            return items[int.Parse(cell.CellValue.Text, System.Globalization.CultureInfo.InvariantCulture)].InnerText;
        }

        return cell.CellValue.Text;
    }

    /// <summary>A part's XML as read, parsed, to compare elements without regard to formatting.</summary>
    public static XElement Xml(OpenXmlPart part)
    {
        using var reader = part.GetStream();
        return XElement.Load(reader);
    }

    public void Dispose()
    {
        Document.Dispose();
        stream.Dispose();
    }
}
