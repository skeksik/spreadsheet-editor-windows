using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>One worksheet of a <see cref="TestWorkbook"/>: cells by A1 reference, and merged ranges.</summary>
internal sealed class TestSheet
{
    private readonly TestWorkbook workbook;
    private readonly SortedDictionary<int, SortedDictionary<int, X.Cell>> rows = [];
    private readonly List<string> merges = [];
    private readonly Dictionary<int, uint> rowStyles = [];
    private readonly List<(uint First, uint Last, uint Style)> columnStyles = [];
    private readonly List<(int Column, int Row, byte[] Bytes, PictureAnchor Anchor)> pictures = [];
    private readonly List<(uint Min, uint Max, double Width)> widths = [];

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
    public TestSheet Empty(string reference, uint style = 1) => Put(reference, new X.Cell(), style);

    /// <summary>Gives a row a style (<c>s</c> with <c>customFormat="1"</c>).</summary>
    public TestSheet RowStyle(int row, uint style)
    {
        rowStyles[row] = style;
        return this;
    }

    /// <summary>Gives columns <paramref name="first"/> to <paramref name="last"/> (1-based) a style with a <c>col</c> element.</summary>
    public TestSheet ColumnStyle(uint first, uint last, uint style)
    {
        columnStyles.Add((first, last, style));
        return this;
    }

    /// <summary>Adds a picture anchored at the cell <paramref name="reference"/>; pictures are written in the order they are added.</summary>
    public TestSheet Picture(string reference, byte[] bytes, PictureAnchor anchor = PictureAnchor.TwoCell)
    {
        var (column, row) = Parse(reference);
        pictures.Add((column, row, bytes, anchor));
        return this;
    }

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

    /// <summary>Gives columns <paramref name="min"/> to <paramref name="max"/> a width (a <c>col</c> element without a style).</summary>
    public TestSheet Width(uint min, uint max, double width)
    {
        widths.Add((min, max, width));
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
        foreach (var index in rowStyles.Keys.Where(r => !rows.ContainsKey(r)))
        {
            rows[index] = [];
        }
        foreach (var (index, cells) in rows)
        {
            var row = new X.Row { RowIndex = (uint)index };
            if (rowStyles.TryGetValue(index, out var rowStyle))
            {
                row.StyleIndex = rowStyle;
                row.CustomFormat = true;
            }
            foreach (var cell in cells.Values)
            {
                row.AppendChild((X.Cell)cell.CloneNode(true));
            }
            data.AppendChild(row);
        }
        var worksheet = new X.Worksheet();
        if (columnStyles.Count > 0 || widths.Count > 0)
        {
            worksheet.AppendChild(new X.Columns(
                columnStyles.Select(c => new X.Column { Min = c.First, Max = c.Last, Style = c.Style, Width = 10, CustomWidth = true })
                    .Concat(widths.Select(c => new X.Column { Min = c.Min, Max = c.Max, Width = c.Width, CustomWidth = true }))));
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

    /// <summary>Writes the pictures into a drawing part of the worksheet part, after <see cref="ToWorksheet"/>.</summary>
    internal void AddDrawing(WorksheetPart part)
    {
        if (pictures.Count == 0)
        {
            return;
        }
        var drawing = part.AddNewPart<DrawingsPart>();
        var root = new Xdr.WorksheetDrawing();
        uint id = 1;
        foreach (var (column, row, bytes, anchor) in pictures)
        {
            string blipId;
            if (anchor == PictureAnchor.Linked)
            {
                blipId = drawing.AddExternalRelationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", new Uri("file:///pictures/linked.png")).Id;
            }
            else
            {
                var image = drawing.AddImagePart(ImagePartType.Png);
                using (var stream = new MemoryStream(bytes))
                {
                    image.FeedData(stream);
                }
                blipId = drawing.GetIdOfPart(image);
            }
            var blip = anchor == PictureAnchor.Linked ? new A.Blip { Link = blipId } : new A.Blip { Embed = blipId };
            var picture = new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = $"Picture {id}" },
                    new Xdr.NonVisualPictureDrawingProperties()),
                new Xdr.BlipFill(blip, new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties());
            id++;
            var extent = new Xdr.Extent { Cx = 914400, Cy = 914400 };
            OpenXmlElement element = anchor switch
            {
                PictureAnchor.OneCell => new Xdr.OneCellAnchor(From(column, row), extent, picture, new Xdr.ClientData()),
                PictureAnchor.Absolute => new Xdr.AbsoluteAnchor(new Xdr.Position { X = 0, Y = 0 }, extent, picture, new Xdr.ClientData()),
                PictureAnchor.Grouped => new Xdr.TwoCellAnchor(
                    From(column, row),
                    To(column, row),
                    new Xdr.GroupShape(
                        new Xdr.NonVisualGroupShapeProperties(
                            new Xdr.NonVisualDrawingProperties { Id = id++, Name = "Group" },
                            new Xdr.NonVisualGroupShapeDrawingProperties()),
                        new Xdr.GroupShapeProperties(),
                        picture),
                    new Xdr.ClientData()),
                _ => new Xdr.TwoCellAnchor(From(column, row), To(column, row), picture, new Xdr.ClientData()),
            };
            root.AppendChild(element);
        }
        drawing.WorksheetDrawing = root;
        part.Worksheet!.AppendChild(new X.Drawing { Id = part.GetIdOfPart(drawing) });
    }

    private static Xdr.FromMarker From(int column, int row) =>
        new(Marker(column - 1), new Xdr.ColumnOffset("0"), new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture)), new Xdr.RowOffset("0"));

    private static Xdr.ToMarker To(int column, int row) =>
        new(Marker(column), new Xdr.ColumnOffset("0"), new Xdr.RowId(row.ToString(CultureInfo.InvariantCulture)), new Xdr.RowOffset("0"));

    private static Xdr.ColumnId Marker(int column) => new(column.ToString(CultureInfo.InvariantCulture));

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
