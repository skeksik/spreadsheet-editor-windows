using SpreadsheetEditor.Core.Imaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Widens a worksheet column for its pictures (docs/spec.md §5.2 step 8): columns are only ever widened.</summary>
internal static class ColumnWidthWriter
{
    /// <summary>
    /// Gives the 1-based <paramref name="column"/> the width <see cref="ColumnWidths.Widen"/> asks for <paramref name="targetPx"/>, when it
    /// is wider than the column's explicit width (or the default width). A <c>col</c> range covering other columns too is split, so that
    /// they keep their settings.
    /// </summary>
    public static void Widen(X.Worksheet worksheet, int column, int targetPx)
    {
        var columns = worksheet.GetFirstChild<X.Columns>();
        var covering = columns?.Elements<X.Column>().FirstOrDefault(c => c.Min?.Value <= column && c.Max?.Value >= column);
        if (ColumnWidths.Widen(covering?.Width?.Value, targetPx) is not { } width)
        {
            return;
        }

        if (covering is null)
        {
            if (columns is null)
            {
                columns = new X.Columns();
                worksheet.InsertBefore(columns, worksheet.GetFirstChild<X.SheetData>());
            }

            var element = new X.Column { Min = (uint)column, Max = (uint)column, Width = width, CustomWidth = true };
            var next = columns.Elements<X.Column>().FirstOrDefault(c => c.Min?.Value > column);
            if (next is null)
            {
                columns.AppendChild(element);
            }
            else
            {
                columns.InsertBefore(element, next);
            }

            return;
        }

        var min = covering.Min!.Value;
        var max = covering.Max!.Value;
        if (min < column)
        {
            var before = (X.Column)covering.CloneNode(true);
            before.Max = (uint)column - 1;
            covering.InsertBeforeSelf(before);
        }

        if (max > column)
        {
            var after = (X.Column)covering.CloneNode(true);
            after.Min = (uint)column + 1;
            covering.InsertAfterSelf(after);
        }

        covering.Min = (uint)column;
        covering.Max = (uint)column;
        covering.Width = width;
        covering.CustomWidth = true;
    }
}
