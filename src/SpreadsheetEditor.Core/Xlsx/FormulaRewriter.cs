using SpreadsheetEditor.Core.Formulas;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Rewrites the formulas inside one worksheet's DOM for a save that renames sheets or moves rows (docs/spec.md §9.20, §9.30), before the
/// rows themselves move: cell formulas, conditional formatting and data validation formulas, hyperlink locations and the formulas of the
/// worksheet's extensions (<c>xm:f</c>).
/// </summary>
/// <remarks>
/// <para>Shared formulas (ECMA-376 Part 1 §18.3.1.40): only the master cell carries text, and every other cell of the group reads it
/// translated by its offset from the master. A group stays shared when, after the rewrite, every remaining cell still reads its master's
/// new text translated by its offset and every cell moved by the same number of rows; then the master's text and <c>ref</c> are rewritten.
/// Otherwise every remaining cell of the group gets its own formula (no new master is chosen), so no cell loses its formula. A cell with
/// text whose <c>si</c> already has a master keeps its own text, as the reader reads it.</para>
/// <para>Array and data table formulas are kept when their whole range moves by one offset; a range that the change splits is refused.</para>
/// </remarks>
internal static class FormulaRewriter
{
    /// <summary>Rewrites the worksheet's cell formulas.</summary>
    /// <param name="worksheet">The worksheet, its rows not moved yet.</param>
    /// <param name="sheetName">The sheet's name, for messages.</param>
    /// <param name="rewrite">The new text of a formula read at a row and column, before the move.</param>
    /// <param name="rowMap">Where each row moves (null for a deleted row), or null when rows do not move.</param>
    public static void RewriteCells(X.Worksheet worksheet, string sheetName, Func<string, int, int, string> rewrite, Func<int, int?>? rowMap)
    {
        var map = rowMap ?? (row => row);
        var masters = new Dictionary<uint, (int Row, int Column, X.CellFormula Formula)>();
        var dependents = new List<(uint Index, int Row, int Column, X.CellFormula Formula)>();
        foreach (var (row, column, cell) in Cells(worksheet))
        {
            if (cell.CellFormula is not { } formula)
            {
                continue;
            }

            var type = formula.FormulaType?.Value;
            var text = formula.Text;
            if (type == X.CellFormulaValues.Shared && formula.SharedIndex?.Value is { } index)
            {
                if (string.IsNullOrEmpty(text))
                {
                    dependents.Add((index, row, column, formula));
                    continue;
                }

                if (masters.TryAdd(index, (row, column, formula)))
                {
                    continue;
                }
            }

            if (type == X.CellFormulaValues.Array || type == X.CellFormulaValues.DataTable)
            {
                MoveRange(formula, row, map, sheetName);
            }

            if (type == X.CellFormulaValues.DataTable)
            {
                formula.R1 = formula.R1?.Value is { Length: > 0 } r1 ? rewrite(r1, row, column) : formula.R1;
                formula.R2 = formula.R2?.Value is { Length: > 0 } r2 ? rewrite(r2, row, column) : formula.R2;
            }

            if (!string.IsNullOrEmpty(text))
            {
                SetText(formula, rewrite(text, row, column));
            }
        }

        foreach (var group in dependents.GroupBy(d => d.Index).ToList())
        {
            if (masters.TryGetValue(group.Key, out var master))
            {
                RewriteGroup(master, [.. group.Select(d => (d.Row, d.Column, d.Formula))], rewrite, map);
            }
        }

        // Masters without dependents still have their own text to rewrite.
        var used = dependents.Select(d => d.Index).ToHashSet();
        foreach (var (index, master) in masters)
        {
            if (!used.Contains(index))
            {
                SetText(master.Formula, rewrite(master.Formula.Text, master.Row, master.Column));
                if (map(master.Row) is { } moved && moved != master.Row)
                {
                    MoveSharedRange(master.Formula, moved - master.Row);
                }
            }
        }
    }

    /// <summary>
    /// Rewrites the worksheet's other formulas, which do not move with a cell: conditional formatting rules, data validations,
    /// hyperlink locations and the extensions' <c>xm:f</c> elements.
    /// </summary>
    public static void RewriteOther(X.Worksheet worksheet, Func<string, string> rewrite)
    {
        foreach (var formula in worksheet.Elements<X.ConditionalFormatting>().SelectMany(c => c.Descendants<X.Formula>()))
        {
            SetLeaf(formula, rewrite);
        }

        foreach (var validation in worksheet.Descendants<X.DataValidation>())
        {
            if (validation.Formula1 is { } first)
            {
                SetLeaf(first, rewrite);
            }

            if (validation.Formula2 is { } second)
            {
                SetLeaf(second, rewrite);
            }
        }

        foreach (var hyperlink in worksheet.Descendants<X.Hyperlink>())
        {
            if (hyperlink.Location?.Value is { Length: > 0 } location)
            {
                var rewritten = rewrite(location);
                if (!string.Equals(rewritten, location, StringComparison.Ordinal))
                {
                    hyperlink.Location = rewritten;
                }
            }
        }

        if (worksheet.GetFirstChild<X.WorksheetExtensionList>() is { } extensions)
        {
            foreach (var formula in extensions.Descendants<Xm.Formula>())
            {
                SetLeaf(formula, rewrite);
            }
        }
    }

    private static void RewriteGroup(
        (int Row, int Column, X.CellFormula Formula) master,
        List<(int Row, int Column, X.CellFormula Formula)> members,
        Func<string, int, int, string> rewrite,
        Func<int, int?> map)
    {
        var source = master.Formula.Text;
        var newMaster = rewrite(source, master.Row, master.Column);
        var masterRow = map(master.Row);
        var texts = new List<(X.CellFormula Formula, string Text)>();
        var shared = masterRow is not null;
        foreach (var (row, column, formula) in members)
        {
            if (map(row) is not { } moved)
            {
                // A deleted cell goes with its row.
                continue;
            }

            var own = rewrite(FormulaReferences.Translate(source, row - master.Row, column - master.Column), row, column);
            texts.Add((formula, own));
            if (shared && (moved - row != masterRow - master.Row
                || !string.Equals(own, FormulaReferences.Translate(newMaster, moved - masterRow!.Value, column - master.Column), StringComparison.Ordinal)))
            {
                shared = false;
            }
        }

        if (shared)
        {
            SetText(master.Formula, newMaster);
            if (masterRow!.Value != master.Row)
            {
                MoveSharedRange(master.Formula, masterRow.Value - master.Row);
            }

            return;
        }

        Unshare(master.Formula, newMaster);
        foreach (var (formula, text) in texts)
        {
            Unshare(formula, text);
        }
    }

    /// <summary>Makes a shared formula's cell a plain formula with this text.</summary>
    public static void Unshare(X.CellFormula formula, string text)
    {
        formula.FormulaType = null;
        formula.SharedIndex = null;
        formula.Reference = null;
        formula.Text = text;
    }

    private static void MoveSharedRange(X.CellFormula formula, int offset)
    {
        if (CellReference.TryParseRange(formula.Reference?.Value, out var range))
        {
            formula.Reference = RowPlan.Name(range with { FirstRow = range.FirstRow + offset, LastRow = range.LastRow + offset });
        }
    }

    /// <summary>Moves an array or data table formula's range with its rows, which must all move by one offset.</summary>
    private static void MoveRange(X.CellFormula formula, int row, Func<int, int?> map, string sheetName)
    {
        if (!CellReference.TryParseRange(formula.Reference?.Value, out var range))
        {
            return;
        }

        var offset = map(range.FirstRow) - range.FirstRow;
        for (var r = range.FirstRow; r <= range.LastRow; r++)
        {
            if (map(r) - r != offset)
            {
                throw SpreadsheetWriteException.NotSupported(
                    $"The array formula in {CellIndex.Name(row, range.FirstColumn)} on the sheet \"{sheetName}\" spans rows that the save moves apart; it cannot be split.");
            }
        }

        if (offset is { } move && move != 0)
        {
            formula.Reference = RowPlan.Name(range with { FirstRow = range.FirstRow + move, LastRow = range.LastRow + move });
        }
    }

    private static void SetText(X.CellFormula formula, string text)
    {
        if (!string.Equals(formula.Text, text, StringComparison.Ordinal))
        {
            formula.Text = text;
        }
    }

    private static void SetLeaf(DocumentFormat.OpenXml.OpenXmlLeafTextElement element, Func<string, string> rewrite)
    {
        var text = element.Text;
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        var rewritten = rewrite(text);
        if (!string.Equals(rewritten, text, StringComparison.Ordinal))
        {
            element.Text = rewritten;
        }
    }

    /// <summary>The worksheet's cells in document order with their positions (rows and cells without <c>r</c> numbered as read).</summary>
    public static IEnumerable<(int Row, int Column, X.Cell Cell)> Cells(X.Worksheet worksheet)
    {
        // Building the index numbers the rows and cells that have no r attribute.
        _ = new CellIndex(worksheet);
        var data = worksheet.GetFirstChild<X.SheetData>();
        if (data is null)
        {
            yield break;
        }

        foreach (var row in data.Elements<X.Row>())
        {
            var rowNumber = (int)row.RowIndex!.Value;
            foreach (var cell in row.Elements<X.Cell>())
            {
                if (CellReference.TryParse(cell.CellReference?.Value, out var column, out _))
                {
                    yield return (rowNumber, column, cell);
                }
            }
        }
    }
}
