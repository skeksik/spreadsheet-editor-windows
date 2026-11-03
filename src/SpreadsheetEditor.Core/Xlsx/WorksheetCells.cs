using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Formulas;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Values;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Reads a worksheet's cells and merged ranges with the SDK's SAX reader (<see cref="OpenXmlReader"/>), so the worksheet's XML is never
/// loaded as a whole. Microsoft Learn, "How to: Parse and read a large spreadsheet document" (read 2026-11-03): "the DOM approach
/// requires loading entire Open XML parts into memory, which can cause an Out of Memory exception … Using the SAX approach, you can
/// employ an OpenXMLReader to read the XML in the file one element at a time". Only non-empty cells are kept, as <see cref="ReadCell"/>s.
/// </summary>
/// <remarks>
/// <para>Shared formulas (ECMA-376 Part 1 §18.3.1.40, <c>t="shared"</c>): the master cell carries the source, a <c>ref</c> and a
/// <c>si</c>; the group's other cells carry only the <c>si</c>. Every cell of a group is shown as its own formula (docs/spec.md §9.3):
/// the master's source translated by the cell's row and column offset from the master (<see cref="FormulaReferences.Translate"/>).
/// The value stays the cell's cached value. Dependents are resolved after the whole worksheet is read, so the master may be anywhere
/// on the sheet, inside the table or not, and before or after them; a dependent whose master is missing keeps its cached value and no
/// formula. The groups are recorded in <see cref="CellGrid.SharedFormulas"/> for the writer.</para>
/// <para>Array formulas (<c>t="array"</c>) and data tables (<c>t="dataTable"</c>) are not translated: the top-left cell, which carries
/// the <c>f</c> element, shows its formula text when it has some (a data table's usually has none), and the other cells of the range,
/// which carry no <c>f</c> element, show their cached value.</para>
/// </remarks>
internal sealed class WorksheetCells
{
    private readonly string[] sharedStrings;
    private readonly NumberFormatTable formats;
    private readonly DateSystem system;
    private readonly CultureInfo culture;

    public WorksheetCells(string[] sharedStrings, NumberFormatTable formats, DateSystem system, CultureInfo culture)
    {
        this.sharedStrings = sharedStrings;
        this.formats = formats;
        this.system = system;
        this.culture = culture;
    }

    public CellGrid Read(WorksheetPart part, CancellationToken cancellationToken)
    {
        var grid = new CellGrid();
        var dependents = new List<PendingCell>();
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
                    Commit(grid, dependents, ref cell);
                    row = int.TryParse(Attribute(reader, "r"), NumberStyles.None, CultureInfo.InvariantCulture, out var r) ? r : row + 1;
                    column = 0;
                    if (Attribute(reader, "customFormat") is "1" or "true" && UInt(Attribute(reader, "s")) is { } rowStyle)
                    {
                        grid.SetRowStyle(row, rowStyle);
                    }
                    if ((++rowsRead & 0xFFF) == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }
                }
                else if (type == typeof(X.Cell))
                {
                    Commit(grid, dependents, ref cell);
                    if (CellReference.TryParse(Attribute(reader, "r"), out var c, out var r))
                    {
                        column = c;
                        row = r;
                    }
                    else
                    {
                        column++;
                    }
                    var explicitStyle = UInt(Attribute(reader, "s"));
                    if (explicitStyle is { } cellStyle)
                    {
                        grid.SetCellStyle(row, column, cellStyle);
                    }
                    cell = new PendingCell(row, column, Attribute(reader, "t"), explicitStyle ?? 0);
                    inPhonetic = false;
                }
                else if (cell is not null && type == typeof(X.CellValue))
                {
                    cell.Raw = reader.GetText();
                }
                else if (cell is not null && type == typeof(X.CellFormula))
                {
                    if (Attribute(reader, "t") == "shared" && UInt(Attribute(reader, "si")) is { } index)
                    {
                        cell.SharedIndex = index;
                        cell.SharedRange = Attribute(reader, "ref");
                    }
                    var formula = reader.GetText();
                    // A shared formula's other cells carry no text: they are given the master's formula once the sheet is read.
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
                else if (type == typeof(X.Column))
                {
                    if (UInt(Attribute(reader, "min")) is { } min && UInt(Attribute(reader, "max")) is { } max && UInt(Attribute(reader, "style")) is { } columnStyle)
                    {
                        grid.SetColumnStyle((int)Math.Min(min, CellReference.MaxColumn), (int)Math.Min(max, CellReference.MaxColumn), columnStyle);
                    }
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
                    Commit(grid, dependents, ref cell);
                }
            }
        }
        Commit(grid, dependents, ref cell);
        ResolveDependents(grid, dependents);
        return grid;
    }

    private void Commit(CellGrid grid, List<PendingCell> dependents, ref PendingCell? cell)
    {
        if (cell is null)
        {
            return;
        }
        var pending = cell;
        cell = null;
        if (pending.SharedIndex is { } index)
        {
            if (pending.Formula is null)
            {
                grid.SharedFormulas.AddCell(pending.Row, pending.Column, new SharedFormulaCell(index, false));
                dependents.Add(pending);
                return;
            }
            // The first cell with text is the master; another cell with text and the same si keeps its own text (not expected).
            var isMaster = grid.SharedFormulas.TryAddGroup(
                new SharedFormulaGroup(index, pending.Row, pending.Column, pending.SharedRange, pending.Formula));
            grid.SharedFormulas.AddCell(pending.Row, pending.Column, new SharedFormulaCell(index, isMaster));
        }
        Add(grid, pending);
    }

    /// <summary>Gives each dependent its master's formula, translated to the dependent's position; without a master, its cached value.</summary>
    private void ResolveDependents(CellGrid grid, List<PendingCell> dependents)
    {
        foreach (var pending in dependents)
        {
            if (grid.SharedFormulas.Group(pending.SharedIndex!.Value) is { } group)
            {
                pending.Formula = FormulaReferences.Translate(group.Source, pending.Row - group.MasterRow, pending.Column - group.MasterColumn);
            }
            Add(grid, pending);
        }
    }

    private void Add(CellGrid grid, PendingCell pending)
    {
        string? text = pending.Type switch
        {
            "s" => SharedString(pending.Raw),
            "inlineStr" => pending.Inline?.ToString() ?? "",
            "str" => pending.Raw ?? "",
            _ => null,
        };
        var formatId = formats.FormatId(pending.Style);
        var format = new NumberFormat(formatId, formats.FormatCode(formatId));
        var value = XlsxCellValues.Create(pending.Type, pending.Raw, text, pending.Formula, format, system, culture);
        if (value is not null)
        {
            grid.Add(new ReadCell(pending.Row, pending.Column, value, text, pending.Formula));
        }
    }

    private string SharedString(string? raw) =>
        int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var index) && index < sharedStrings.Length
            ? sharedStrings[index]
            : throw new InvalidDataException($"A cell refers to shared string {raw}, which the workbook does not have.");

    private static uint? UInt(string? text) =>
        uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var value) ? value : null;

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

        /// <summary>The <c>si</c> of a shared formula's <c>f</c> element, or null when the cell is not in a shared formula group.</summary>
        public uint? SharedIndex { get; set; }

        /// <summary>A shared formula master's <c>ref</c>.</summary>
        public string? SharedRange { get; set; }

        public StringBuilder? Inline { get; set; }
    }
}
