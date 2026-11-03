using System.Globalization;
using DocumentFormat.OpenXml;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>One worksheet of a <see cref="TestWorkbook"/>: cells by A1 reference, and merged ranges.</summary>
internal sealed class TestSheet
{
    private readonly TestWorkbook workbook;
    private readonly SortedDictionary<int, SortedDictionary<int, X.Cell>> rows = [];
    private readonly List<string> merges = [];

    internal TestSheet(TestWorkbook workbook) => this.workbook = workbook;

    /// <summary>Sets cells along a row from <paramref name="reference"/>: strings (shared), numbers and booleans; null skips a cell.</summary>
    public TestSheet Row(string reference, params object?[] values)
    {
        var (column, row) = Parse(reference);
        foreach (var value in values)
        {
            if (value is not null)
            {
                Set(Name(column) + row.ToString(CultureInfo.InvariantCulture), value);
            }
            column++;
        }
        return this;
    }

    /// <summary>Sets one cell: a string becomes a shared string, a number a number, a bool a boolean.</summary>
    public TestSheet Set(string reference, object value, uint style = 0)
    {
        var cell = value switch
        {
            string text => new X.Cell
            {
                DataType = X.CellValues.SharedString,
                CellValue = new X.CellValue(workbook.AddSharedString(new X.SharedStringItem(new X.Text(text) { Space = SpaceProcessingModeValues.Preserve })).ToString(CultureInfo.InvariantCulture)),
            },
            bool flag => new X.Cell { DataType = X.CellValues.Boolean, CellValue = new X.CellValue(flag ? "1" : "0") },
            _ => new X.Cell { CellValue = new X.CellValue(Convert.ToString(value, CultureInfo.InvariantCulture)!) },
        };
        return Put(reference, cell, style);
    }

    /// <summary>Sets a cell with an inline string.</summary>
    public TestSheet Inline(string reference, string text) =>
        Put(reference, new X.Cell { DataType = X.CellValues.InlineString, InlineString = new X.InlineString(new X.Text(text)) }, 0);

    /// <summary>Sets a cell to a shared string made of these runs, with a phonetic run that is not part of the text.</summary>
    public TestSheet RichText(string reference, params string[] runs)
    {
        var item = new X.SharedStringItem(runs.Select(r => new X.Run(new X.Text(r))));
        item.AppendChild(new X.PhoneticRun(new X.Text("phonetic")) { BaseTextStartIndex = 0, EndingBaseIndex = 1 });
        var index = workbook.AddSharedString(item);
        return Put(reference, new X.Cell { DataType = X.CellValues.SharedString, CellValue = new X.CellValue(index.ToString(CultureInfo.InvariantCulture)) }, 0);
    }

    /// <summary>Sets a cell with a formula and its cached value (a number, a string or a bool).</summary>
    public TestSheet Formula(string reference, string formula, object cached)
    {
        var cell = cached switch
        {
            string text => new X.Cell { DataType = X.CellValues.String, CellValue = new X.CellValue(text) },
            bool flag => new X.Cell { DataType = X.CellValues.Boolean, CellValue = new X.CellValue(flag ? "1" : "0") },
            _ => new X.Cell { CellValue = new X.CellValue(Convert.ToString(cached, CultureInfo.InvariantCulture)!) },
        };
        cell.CellFormula = new X.CellFormula(formula);
        return Put(reference, cell, 0);
    }

    /// <summary>Sets an error cell, such as <c>#N/A</c>.</summary>
    public TestSheet Error(string reference, string error) =>
        Put(reference, new X.Cell { DataType = X.CellValues.Error, CellValue = new X.CellValue(error) }, 0);

    /// <summary>Adds a cell with a style and no value.</summary>
    public TestSheet Empty(string reference) => Put(reference, new X.Cell(), 1);

    public TestSheet Merge(string range)
    {
        merges.Add(range);
        return this;
    }

    internal X.Worksheet ToWorksheet()
    {
        var data = new X.SheetData();
        foreach (var (index, cells) in rows)
        {
            var row = new X.Row { RowIndex = (uint)index };
            foreach (var cell in cells.Values)
            {
                row.AppendChild((X.Cell)cell.CloneNode(true));
            }
            data.AppendChild(row);
        }
        var worksheet = new X.Worksheet(data);
        if (merges.Count > 0)
        {
            worksheet.AppendChild(new X.MergeCells(merges.Select(m => new X.MergeCell { Reference = m })));
        }
        return worksheet;
    }

    private TestSheet Put(string reference, X.Cell cell, uint style)
    {
        var (column, row) = Parse(reference);
        cell.CellReference = reference;
        if (style != 0)
        {
            cell.StyleIndex = style;
        }
        if (!rows.TryGetValue(row, out var cells))
        {
            cells = [];
            rows[row] = cells;
        }
        cells[column] = cell;
        return this;
    }

    private static (int Column, int Row) Parse(string reference)
    {
        var letters = reference.TakeWhile(char.IsAsciiLetter).Count();
        var column = 0;
        foreach (var c in reference[..letters])
        {
            column = (column * 26) + (char.ToUpperInvariant(c) - 'A' + 1);
        }
        return (column, int.Parse(reference[letters..], CultureInfo.InvariantCulture));
    }

    private static string Name(int column)
    {
        var name = "";
        while (column > 0)
        {
            name = (char)('A' + ((column - 1) % 26)) + name;
            column = (column - 1) / 26;
        }
        return name;
    }
}
