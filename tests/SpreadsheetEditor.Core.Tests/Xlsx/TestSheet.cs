using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>One worksheet of a <see cref="TestWorkbook"/>: cells by A1 reference, and merged ranges.</summary>
internal sealed class TestSheet
{
    private readonly TestWorkbook workbook;
    private readonly SortedDictionary<int, SortedDictionary<int, X.Cell>> rows = [];
    private readonly List<string> merges = [];
    private readonly List<(uint Min, uint Max, double Width)> columns = [];
    private readonly List<string> pictures = [];

    /// <summary>When set, the worksheet has page margins and a landscape page setup.</summary>
    public bool PageSetup { get; set; }

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

    /// <summary>Sets a shared formula's first cell, whose formula is used by every cell of <paramref name="range"/>.</summary>
    public TestSheet SharedFormula(string reference, string formula, string range, double cached)
    {
        var cell = new X.Cell { CellValue = new X.CellValue(cached.ToString(CultureInfo.InvariantCulture)) };
        cell.CellFormula = new X.CellFormula(formula) { FormulaType = X.CellFormulaValues.Shared, Reference = range, SharedIndex = 0 };
        return Put(reference, cell, 0);
    }

    /// <summary>Sets another cell of shared formula 0: no formula text, only its cached value.</summary>
    public TestSheet SharedDependent(string reference, double cached)
    {
        var cell = new X.Cell { CellValue = new X.CellValue(cached.ToString(CultureInfo.InvariantCulture)) };
        cell.CellFormula = new X.CellFormula { FormulaType = X.CellFormulaValues.Shared, SharedIndex = 0 };
        return Put(reference, cell, 0);
    }

    /// <summary>Sets a cell with a style and a string.</summary>
    public TestSheet Styled(string reference, string text, uint style) => Set(reference, text, style);

    /// <summary>Gives columns <paramref name="min"/> to <paramref name="max"/> a width.</summary>
    public TestSheet Width(uint min, uint max, double width)
    {
        columns.Add((min, max, width));
        return this;
    }

    /// <summary>Adds a picture (a 1×1 PNG) anchored at the cell, in a drawing part.</summary>
    public TestSheet Picture(string reference)
    {
        pictures.Add(reference);
        return this;
    }

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
        var worksheet = new X.Worksheet();
        if (columns.Count > 0)
        {
            worksheet.AppendChild(new X.Columns(columns.Select(c => new X.Column { Min = c.Min, Max = c.Max, Width = c.Width, CustomWidth = true })));
        }
        worksheet.AppendChild(data);
        if (merges.Count > 0)
        {
            worksheet.AppendChild(new X.MergeCells(merges.Select(m => new X.MergeCell { Reference = m })));
        }
        if (PageSetup)
        {
            worksheet.AppendChild(new X.PageMargins { Left = 0.5, Right = 0.5, Top = 0.75, Bottom = 0.75, Header = 0.3, Footer = 0.3 });
            worksheet.AppendChild(new X.PageSetup { Orientation = X.OrientationValues.Landscape });
        }
        return worksheet;
    }

    /// <summary>Adds the drawing part with the pictures, and the worksheet's reference to it.</summary>
    internal void AddParts(WorksheetPart part)
    {
        if (pictures.Count == 0)
        {
            return;
        }

        var drawing = part.AddNewPart<DrawingsPart>();
        var image = drawing.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(Png))
        {
            image.FeedData(stream);
        }

        var imageId = drawing.GetIdOfPart(image);
        var anchors = new System.Text.StringBuilder();
        var id = 1;
        foreach (var reference in pictures)
        {
            var (column, row) = Parse(reference);
            anchors.Append(CultureInfo.InvariantCulture, $"""
                <xdr:oneCellAnchor><xdr:from><xdr:col>{column - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="9525" cy="9525"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id}" name="Image {id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="{imageId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
                """);
            id++;
        }

        var xml = $"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{anchors}</xdr:wsDr>""";
        using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml)))
        {
            drawing.FeedData(stream);
        }

        part.Worksheet!.AppendChild(new X.Drawing { Id = part.GetIdOfPart(drawing) });
    }

    /// <summary>A 1×1 transparent PNG.</summary>
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

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
