using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Reads an <c>.xlsx</c> file into a <see cref="Workbook"/> (docs/spec.md §4.1–4.5). The whole file is parsed before a workbook is
/// returned; on any failure a <see cref="SpreadsheetReadException"/> is thrown and nothing is returned (§9.13). Reading does not touch
/// the UI: call <see cref="ReadAsync"/>, or <see cref="Read(string, CultureInfo, CancellationToken)"/> from a worker thread.
/// </summary>
public static class XlsxReader
{
    private static ReadOnlySpan<byte> ZipSignature => [0x50, 0x4B, 0x03, 0x04];

    private static ReadOnlySpan<byte> EmptyZipSignature => [0x50, 0x4B, 0x05, 0x06];

    private static ReadOnlySpan<byte> CompoundFileSignature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    /// <summary>Reads the file on the thread pool.</summary>
    /// <param name="path">The file's path.</param>
    /// <param name="culture">The UI language: values are shown as text in its formats (dates, numbers).</param>
    /// <param name="cancellationToken">Stops the read; the task is then cancelled.</param>
    public static Task<Workbook> ReadAsync(string path, CultureInfo culture, CancellationToken cancellationToken = default) =>
        Task.Run(() => Read(path, culture, cancellationToken), cancellationToken);

    /// <summary>Reads the file. The file is read into memory once and not kept open or locked.</summary>
    /// <param name="path">The file's path; it becomes <see cref="Workbook.SourcePath"/>.</param>
    /// <param name="culture">The UI language: values are shown as text in its formats (dates, numbers).</param>
    /// <param name="cancellationToken">Stops the read with an <see cref="OperationCanceledException"/>.</param>
    public static Workbook Read(string path, CultureInfo culture, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(culture);
        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(path);
        }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException)
        {
            throw new SpreadsheetReadException(SpreadsheetReadError.NotFound, $"The file was not found: {ex.Message}", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new SpreadsheetReadException(SpreadsheetReadError.AccessDenied, $"The file cannot be opened: {ex.Message}", ex);
        }
        catch (IOException ex)
        {
            throw new SpreadsheetReadException(SpreadsheetReadError.ReadFailed, $"The file could not be read: {ex.Message}", ex);
        }
        var workbook = Parse(bytes, culture, cancellationToken);
        workbook.SourcePath = path;
        return workbook;
    }

    /// <summary>Reads a workbook from a stream, which is read to its end and not closed.</summary>
    /// <param name="stream">The stream, read from its current position.</param>
    /// <param name="culture">The UI language: values are shown as text in its formats (dates, numbers).</param>
    /// <param name="cancellationToken">Stops the read with an <see cref="OperationCanceledException"/>.</param>
    public static Workbook Read(Stream stream, CultureInfo culture, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(culture);
        byte[] bytes;
        try
        {
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            bytes = copy.ToArray();
        }
        catch (Exception ex) when (ex is IOException or NotSupportedException or ObjectDisposedException)
        {
            throw new SpreadsheetReadException(SpreadsheetReadError.ReadFailed, $"The file could not be read: {ex.Message}", ex);
        }
        return Parse(bytes, culture, cancellationToken);
    }

    private static Workbook Parse(byte[] bytes, CultureInfo culture, CancellationToken cancellationToken)
    {
        CheckSignature(bytes);
        List<Sheet> sheets;
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var document = SpreadsheetDocument.Open(stream, false);
            sheets = ReadSheets(document, culture, cancellationToken);
        }
        catch (Exception ex) when (ex is not (SpreadsheetReadException or OperationCanceledException or OutOfMemoryException))
        {
            // The package or its XML is damaged, or holds something the reader does not expect: report it, never crash (CLAUDE.md).
            throw new SpreadsheetReadException(SpreadsheetReadError.Corrupt, $"The workbook is damaged and cannot be read: {ex.Message}", ex);
        }
        var workbook = new Workbook { Source = new XlsxSource(bytes) };
        workbook.Sheets.AddRange(sheets);
        return workbook;
    }

    private static void CheckSignature(byte[] bytes)
    {
        var span = bytes.AsSpan();
        if (span.StartsWith(ZipSignature) || span.StartsWith(EmptyZipSignature))
        {
            return;
        }
        if (span.StartsWith(CompoundFileSignature))
        {
            throw new SpreadsheetReadException(
                SpreadsheetReadError.LegacyOrEncrypted,
                "The file is an older Excel workbook (.xls) or a password-protected one; only unprotected .xlsx files can be opened.");
        }
        throw new SpreadsheetReadException(SpreadsheetReadError.NotAnXlsx, "The file is not an Excel workbook (.xlsx).");
    }

    private static List<Sheet> ReadSheets(SpreadsheetDocument document, CultureInfo culture, CancellationToken cancellationToken)
    {
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("The package has no workbook part.");
        var root = workbookPart.Workbook
            ?? throw new InvalidDataException("The workbook part is empty.");
        var date1904 = root.WorkbookProperties?.Date1904?.Value ?? false;
        var formats = NumberFormatTable.Read(workbookPart.WorkbookStylesPart?.Stylesheet);
        var sharedStrings = SharedStrings.Read(workbookPart.SharedStringTablePart, cancellationToken);
        var cells = new WorksheetCells(sharedStrings, formats, date1904, culture);

        var sheets = new List<Sheet>();
        foreach (var entry in root.Sheets?.Elements<X.Sheet>() ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            var id = entry.Id?.Value;
            // Chart sheets, dialog sheets and macro sheets are not grid worksheets: skipped, and kept as they are in the package (§9.12).
            if (id is null || FindPart(workbookPart, id) is not WorksheetPart part)
            {
                continue;
            }
            var grid = cells.Read(part, cancellationToken);
            var sheet = SheetBuilder.Build(entry.Name?.Value ?? "", id, entry.SheetId?.Value ?? 0, grid);
            if (sheet is null)
            {
                // No occupied cell: no tab, and the worksheet stays untouched in the package (§9.10).
                continue;
            }
            sheets.Add(sheet);
        }
        return sheets;
    }

    private static OpenXmlPart? FindPart(WorkbookPart workbookPart, string relationshipId)
    {
        foreach (var pair in workbookPart.Parts)
        {
            if (pair.RelationshipId == relationshipId)
            {
                return pair.OpenXmlPart;
            }
        }
        return null;
    }
}
