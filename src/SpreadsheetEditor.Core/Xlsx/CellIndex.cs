using System.Globalization;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The cells of a worksheet's DOM by position, for the writer: finds a cell, and creates it (and its row) in document order when it is
/// missing. Rows and cells without an <c>r</c> attribute get one, numbered as the reader numbers them (the next row or column).
/// </summary>
internal sealed class CellIndex
{
    private readonly X.SheetData data;
    private readonly SortedDictionary<int, X.Row> rows = [];
    private readonly Dictionary<long, X.Cell> cells = [];

    public CellIndex(X.Worksheet worksheet)
    {
        data = worksheet.GetFirstChild<X.SheetData>() ?? worksheet.AppendChild(new X.SheetData());
        var rowNumber = 0;
        foreach (var row in data.Elements<X.Row>())
        {
            rowNumber = row.RowIndex?.Value is { } r ? (int)r : rowNumber + 1;
            row.RowIndex ??= (uint)rowNumber;
            rows[rowNumber] = row;
            var column = 0;
            foreach (var cell in row.Elements<X.Cell>())
            {
                if (CellReference.TryParse(cell.CellReference?.Value, out var c, out _))
                {
                    column = c;
                }
                else
                {
                    column++;
                    cell.CellReference = Name(rowNumber, column);
                }

                cells[Key(rowNumber, column)] = cell;
            }
        }
    }

    public IEnumerable<(int Row, int Column, X.Cell Cell)> All =>
        cells.Select(pair => ((int)(pair.Key >> 16), (int)(pair.Key & 0xFFFF), pair.Value));

    public X.Cell? Get(int row, int column) => cells.GetValueOrDefault(Key(row, column));

    public X.Cell GetOrCreate(int row, int column)
    {
        if (cells.TryGetValue(Key(row, column), out var existing))
        {
            return existing;
        }

        if (!rows.TryGetValue(row, out var rowElement))
        {
            rowElement = new X.Row { RowIndex = (uint)row };
            var next = rows.FirstOrDefault(pair => pair.Key > row).Value;
            if (next is null)
            {
                data.AppendChild(rowElement);
            }
            else
            {
                data.InsertBefore(rowElement, next);
            }

            rows[row] = rowElement;
        }

        var cell = new X.Cell { CellReference = Name(row, column) };
        X.Cell? after = null;
        foreach (var other in rowElement.Elements<X.Cell>())
        {
            if (CellReference.TryParse(other.CellReference?.Value, out var c, out _) && c > column)
            {
                after = other;
                break;
            }
        }

        if (after is null)
        {
            rowElement.AppendChild(cell);
        }
        else
        {
            rowElement.InsertBefore(cell, after);
        }

        cells[Key(row, column)] = cell;
        return cell;
    }

    public static string Name(int row, int column) => CellReference.ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

    private static long Key(int row, int column) => ((long)row << 16) | (uint)column;
}
