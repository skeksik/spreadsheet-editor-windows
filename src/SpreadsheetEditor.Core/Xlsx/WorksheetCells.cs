using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Reads a worksheet's cells and merged ranges with the SDK's SAX reader (<see cref="OpenXmlReader"/>), so the worksheet's XML is never
/// loaded as a whole. Microsoft Learn, "How to: Parse and read a large spreadsheet document" (read 2026-11-03): "the DOM approach
/// requires loading entire Open XML parts into memory, which can cause an Out of Memory exception … Using the SAX approach, you can
/// employ an OpenXMLReader to read the XML in the file one element at a time". Only non-empty cells are kept, as <see cref="ReadCell"/>s.
/// </summary>
internal sealed class WorksheetCells
{
    private readonly string[] sharedStrings;
    private readonly NumberFormatTable formats;
    private readonly bool date1904;
    private readonly CultureInfo culture;

    public WorksheetCells(string[] sharedStrings, NumberFormatTable formats, bool date1904, CultureInfo culture)
    {
        this.sharedStrings = sharedStrings;
        this.formats = formats;
        this.date1904 = date1904;
        this.culture = culture;
    }

    public CellGrid Read(WorksheetPart part, CancellationToken cancellationToken)
    {
        var grid = new CellGrid();
        var row = 0;
        var column = 0;
        var rowsRead = 0;
        PendingCell? cell = null;
        var inPhonetic = false;
        using var reader = OpenXmlReader.Create(part);
        while (reader.Read())
        {
            var type = reader.ElementType;
            if (reader.IsStartElement)
            {
                if (type == typeof(X.Row))
                {
                    Commit(grid, ref cell);
                    row = int.TryParse(Attribute(reader, "r"), NumberStyles.None, CultureInfo.InvariantCulture, out var r) ? r : row + 1;
                    column = 0;
                    if ((++rowsRead & 0xFFF) == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }
                }
                else if (type == typeof(X.Cell))
                {
                    Commit(grid, ref cell);
                    if (CellReference.TryParse(Attribute(reader, "r"), out var c, out var r))
                    {
                        column = c;
                        row = r;
                    }
                    else
                    {
                        column++;
                    }
                    var style = uint.TryParse(Attribute(reader, "s"), NumberStyles.None, CultureInfo.InvariantCulture, out var s) ? s : 0;
                    cell = new PendingCell(row, column, Attribute(reader, "t"), style);
                    inPhonetic = false;
                }
                else if (cell is not null && type == typeof(X.CellValue))
                {
                    cell.Raw = reader.GetText();
                }
                else if (cell is not null && type == typeof(X.CellFormula))
                {
                    var formula = reader.GetText();
                    // A shared formula's other cells carry no text; until it is translated to each cell they keep their cached value.
                    cell.Formula = string.IsNullOrEmpty(formula) ? null : formula;
                }
                else if (cell is not null && type == typeof(X.PhoneticRun))
                {
                    inPhonetic = true;
                }
                else if (cell is not null && type == typeof(X.Run))
                {
                    inPhonetic = false;
                }
                else if (cell is not null && type == typeof(X.Text) && !inPhonetic)
                {
                    (cell.Inline ??= new StringBuilder()).Append(reader.GetText());
                }
                else if (type == typeof(X.MergeCell))
                {
                    if (CellReference.TryParseRange(Attribute(reader, "ref"), out var range))
                    {
                        grid.Merges.Add(range);
                    }
                }
            }
            else if (reader.IsEndElement)
            {
                if (type == typeof(X.PhoneticRun))
                {
                    inPhonetic = false;
                }
                else if (type == typeof(X.Cell) || type == typeof(X.SheetData))
                {
                    Commit(grid, ref cell);
                }
            }
        }
        Commit(grid, ref cell);
        return grid;
    }

    private void Commit(CellGrid grid, ref PendingCell? cell)
    {
        if (cell is null)
        {
            return;
        }
        var pending = cell;
        cell = null;
        string? text = pending.Type switch
        {
            "s" => SharedString(pending.Raw),
            "inlineStr" => pending.Inline?.ToString() ?? "",
            "str" => pending.Raw ?? "",
            _ => null,
        };
        var formatId = formats.FormatId(pending.Style);
        var value = XlsxCellValues.Create(
            pending.Type, pending.Raw, text, pending.Formula, formatId, formats.FormatCode(formatId), date1904, culture);
        if (value is not null)
        {
            grid.Add(new ReadCell(pending.Row, pending.Column, value, text, pending.Formula));
        }
    }

    private string SharedString(string? raw) =>
        int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var index) && index < sharedStrings.Length
            ? sharedStrings[index]
            : throw new InvalidDataException($"A cell refers to shared string {raw}, which the workbook does not have.");

    private static string? Attribute(OpenXmlReader reader, string localName)
    {
        foreach (var attribute in reader.Attributes)
        {
            if (attribute.LocalName == localName && string.IsNullOrEmpty(attribute.NamespaceUri))
            {
                return attribute.Value;
            }
        }
        return null;
    }

    private sealed class PendingCell(int row, int column, string? type, uint style)
    {
        public int Row { get; } = row;

        public int Column { get; } = column;

        public string? Type { get; } = type;

        public uint Style { get; } = style;

        public string? Raw { get; set; }

        public string? Formula { get; set; }

        public StringBuilder? Inline { get; set; }
    }
}
