using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>An empty <c>.xlsx</c> package, the start point for a workbook that was not read from one (docs/spec.md §5.1).</summary>
internal static class NewPackage
{
    /// <summary>
    /// A package with a workbook part without sheets, a style sheet with the default entries and an empty shared string table. It has no
    /// default worksheet: the writer appends one per sheet.
    /// </summary>
    public static byte[] Create()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddNewPart<WorkbookPart>(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "rId1");
            workbookPart.Workbook = new X.Workbook(new X.Sheets());
            workbookPart.AddNewPart<WorkbookStylesPart>("rId1").Stylesheet = DefaultStylesheet();
            workbookPart.AddNewPart<SharedStringTablePart>("rId2").SharedStringTable = new X.SharedStringTable { Count = 0, UniqueCount = 0 };
        }

        return stream.ToArray();
    }

    /// <summary>
    /// The style sheet of a new workbook: one font (Calibri 11), the two fills every workbook starts with (none and gray125), one border,
    /// one cell style format, one cell format and the Normal cell style. That a spreadsheet program expects exactly these first entries is
    /// not verified.
    /// </summary>
    public static X.Stylesheet DefaultStylesheet() => new(
        new X.Fonts(new X.Font(new X.FontSize { Val = 11 }, new X.FontName { Val = "Calibri" }, new X.FontFamilyNumbering { Val = 2 }))
        { Count = 1 },
        new X.Fills(
            new X.Fill(new X.PatternFill { PatternType = X.PatternValues.None }),
            new X.Fill(new X.PatternFill { PatternType = X.PatternValues.Gray125 }))
        { Count = 2 },
        new X.Borders(new X.Border(new X.LeftBorder(), new X.RightBorder(), new X.TopBorder(), new X.BottomBorder(), new X.DiagonalBorder()))
        { Count = 1 },
        new X.CellStyleFormats(new X.CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }) { Count = 1 },
        new X.CellFormats(new X.CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 }) { Count = 1 },
        new X.CellStyles(new X.CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 });
}
