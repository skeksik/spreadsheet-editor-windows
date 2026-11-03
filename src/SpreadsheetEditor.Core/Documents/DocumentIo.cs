using System.Globalization;
using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Editing;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Documents;

/// <summary>
/// Opens, saves and creates documents for the app's File menu (docs/spec.md §2.1-2.2 as replaced by §10.4-10.5): the file's format
/// decides the reader and writer, and the work runs on the thread pool (§9.53).
/// </summary>
public static class DocumentIo
{
    /// <summary>
    /// True when <paramref name="path"/> names a file the app opens: <c>.xlsx</c> or <c>.csv</c>, ignoring case. <c>.xls</c> is not one
    /// (§9.11: not offered by the open picker or accepted on drop).
    /// </summary>
    public static bool CanOpen(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) || extension.Equals(".csv", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Reads the file at <paramref name="path"/> on the thread pool: <c>.csv</c> with <see cref="CsvReader"/>, anything else with
    /// <see cref="XlsxReader"/> (which says plainly when it is not a workbook). Nothing changes on failure, so the open document stays
    /// as it was (§9.13).
    /// </summary>
    /// <param name="path">The file.</param>
    /// <param name="culture">The UI language values are shown in.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <exception cref="SpreadsheetReadException">The file cannot be read, including a <c>.csv</c> that is missing, locked or not permitted.</exception>
    /// <exception cref="CsvReadException">The <c>.csv</c> file is empty, has an unsupported encoding or an unclosed quote.</exception>
    public static async Task<OpenedDocument> OpenAsync(string path, CultureInfo culture, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(culture);
        if (FileFormats.FromPath(path, FileFormat.Xlsx) == FileFormat.Csv)
        {
            var result = await Task.Run(() => ReadCsv(path), cancellationToken).ConfigureAwait(false);
            return new OpenedDocument(result.Workbook, new DocumentFile(path, FileFormat.Csv, result.Format));
        }
        var workbook = await XlsxReader.ReadAsync(path, culture, cancellationToken).ConfigureAwait(false);
        return new OpenedDocument(workbook, new DocumentFile(path, FileFormat.Xlsx, null));
    }

    /// <summary>
    /// Writes the document to <paramref name="path"/> on the thread pool, through a temporary file that replaces the target, so a failed
    /// save leaves the file as it was. <c>.xlsx</c> writes every sheet; <c>.csv</c> writes <paramref name="currentSheet"/> only, in
    /// <see cref="DocumentFile.CsvFormatForSave"/> (§9.15, §9.34).
    /// </summary>
    /// <param name="workbook">The document's workbook.</param>
    /// <param name="currentSheet">The selected sheet, which a <c>.csv</c> save writes; ignored for <c>.xlsx</c>.</param>
    /// <param name="file">The document's file before the save; its read CSV format is kept.</param>
    /// <param name="path">The file to write.</param>
    /// <param name="format">The format to write, from <see cref="FileFormats.FromPath"/> (§9.35).</param>
    /// <param name="culture">The UI language values were shown and edited in.</param>
    /// <param name="cancellationToken">Cancels the save before the target is replaced.</param>
    /// <exception cref="SpreadsheetWriteException">The file cannot be written, or the writer does not support a change (<see cref="SpreadsheetWriteError.NotSupported"/>).</exception>
    public static async Task<SaveResult> SaveAsync(
        Workbook workbook,
        Sheet? currentSheet,
        DocumentFile file,
        string path,
        FileFormat format,
        CultureInfo culture,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(file);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(culture);
        var saved = file.SavedAs(path, format);
        if (format == FileFormat.Csv)
        {
            var sheet = currentSheet ?? workbook.Sheets.FirstOrDefault()
                ?? throw new SpreadsheetWriteException(SpreadsheetWriteError.InvalidWorkbook, "The workbook has no sheet to save.");
            if (!workbook.Sheets.Contains(sheet))
            {
                throw new ArgumentException("The sheet is not in the workbook.", nameof(currentSheet));
            }
            await CsvWriter.WriteAsync(sheet, path, file.CsvFormatForSave, cancellationToken).ConfigureAwait(false);
            return new SaveResult(saved, workbook.Sheets.Count == 1);
        }
        await XlsxWriter.WriteAsync(workbook, path, culture, cancellationToken).ConfigureAwait(false);
        return new SaveResult(saved, true);
    }

    /// <summary>
    /// A new document (§10.5, New): one sheet with the defaults of §2.3, named <paramref name="sheetBaseName"/> followed by 1 under the
    /// sheet-name rules (<see cref="SheetNames.NewSheetName"/>), for example "Sheet1".
    /// </summary>
    /// <param name="sheetBaseName">The app's localised word for a sheet.</param>
    public static OpenedDocument New(string sheetBaseName)
    {
        ArgumentNullException.ThrowIfNull(sheetBaseName);
        var editor = new WorkbookEditor(new Workbook());
        editor.AddSheet(sheetBaseName);
        return new OpenedDocument(editor.Workbook, DocumentFile.Untitled);
    }

    // CsvReader lets file errors through as they are; they become the same read errors an .xlsx gives, so the app has one set of messages.
    private static CsvReadResult ReadCsv(string path)
    {
        try
        {
            return CsvReader.Read(path);
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
    }
}
