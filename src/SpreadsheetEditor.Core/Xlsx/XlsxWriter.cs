using System.Globalization;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Saves a <see cref="Workbook"/> as <c>.xlsx</c> (docs/spec.md §5). A workbook read from <c>.xlsx</c> is written into a fresh copy of that
/// file, so whatever the app does not edit is written back as it was read and two saves in a row give the same file (§9.21); any other
/// workbook starts from a new empty package. The whole file is built in memory before anything is written: on any failure a
/// <see cref="SpreadsheetWriteException"/> is thrown and the target is unchanged. Writing does not touch the UI: call
/// <see cref="WriteAsync"/>, or <see cref="Write(Workbook, string, CultureInfo, CancellationToken)"/> from a worker thread. The workbook
/// must not be edited while it is written (ADR 0002); the model is not changed by a save.
/// </summary>
public static class XlsxWriter
{
    /// <summary>Writes the file on the thread pool.</summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="path">The file to write; it is replaced when it exists.</param>
    /// <param name="culture">The UI language the values were shown and edited in: edited text is parsed in its formats.</param>
    /// <param name="cancellationToken">Stops the write before the target is replaced; the task is then cancelled.</param>
    public static Task WriteAsync(Workbook workbook, string path, CultureInfo culture, CancellationToken cancellationToken = default) =>
        Task.Run(() => Write(workbook, path, culture, cancellationToken), cancellationToken);

    /// <summary>
    /// Writes the file: into a temporary file in the target's folder, which then replaces the target, so the target is never left half
    /// written. (That a replace by <see cref="File.Move(string, string, bool)"/> on the same volume is atomic is not verified.)
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="path">The file to write; it is replaced when it exists.</param>
    /// <param name="culture">The UI language the values were shown and edited in: edited text is parsed in its formats.</param>
    /// <param name="cancellationToken">Stops the write with an <see cref="OperationCanceledException"/> before the target is replaced.</param>
    public static void Write(Workbook workbook, string path, CultureInfo culture, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(culture);

        var bytes = Save(workbook, culture, cancellationToken);
        string? temporary = null;
        try
        {
            var full = Path.GetFullPath(path);
            var folder = Path.GetDirectoryName(full) ?? ".";
            temporary = Path.Combine(folder, $"{Path.GetFileName(full)}.{Guid.NewGuid():N}.tmp");
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(bytes);
                stream.Flush(flushToDisk: true);
            }

            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporary, full, overwrite: true);
            temporary = null;
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new SpreadsheetWriteException(SpreadsheetWriteError.AccessDenied, $"The file cannot be written: {ex.Message}", ex);
        }
        catch (IOException ex)
        {
            throw new SpreadsheetWriteException(SpreadsheetWriteError.WriteFailed, $"The file could not be written: {ex.Message}", ex);
        }
        finally
        {
            if (temporary is not null)
            {
                TryDelete(temporary);
            }
        }
    }

    /// <summary>Writes the workbook to a stream, from its current position; the stream is not closed.</summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="stream">The stream; nothing is written to it when the workbook cannot be written.</param>
    /// <param name="culture">The UI language the values were shown and edited in: edited text is parsed in its formats.</param>
    /// <param name="cancellationToken">Stops the write with an <see cref="OperationCanceledException"/> before the stream is written.</param>
    public static void Write(Workbook workbook, Stream stream, CultureInfo culture, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(culture);

        var bytes = Save(workbook, culture, cancellationToken);
        try
        {
            stream.Write(bytes);
        }
        catch (Exception ex) when (ex is IOException or NotSupportedException or ObjectDisposedException)
        {
            throw new SpreadsheetWriteException(SpreadsheetWriteError.WriteFailed, $"The file could not be written: {ex.Message}", ex);
        }
    }

    /// <summary>The <c>.xlsx</c> file's bytes.</summary>
    internal static byte[] Save(Workbook workbook, CultureInfo culture, CancellationToken cancellationToken)
    {
        try
        {
            return WorkbookWriter.Write(workbook, culture, cancellationToken);
        }
        catch (Exception ex) when (ex is not (SpreadsheetWriteException or OperationCanceledException or OutOfMemoryException))
        {
            // The source holds something the writer does not expect: report it, never crash (CLAUDE.md).
            throw new SpreadsheetWriteException(SpreadsheetWriteError.WriteFailed, $"The workbook could not be written: {ex.Message}", ex);
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A temporary file that cannot be removed is left behind; the target is unchanged either way.
        }
    }
}
