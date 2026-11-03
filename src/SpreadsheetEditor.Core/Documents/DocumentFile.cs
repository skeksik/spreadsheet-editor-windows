using SpreadsheetEditor.Core.Csv;

namespace SpreadsheetEditor.Core.Documents;

/// <summary>
/// The file an open document belongs to (docs/spec.md §10.5): where Save writes and in which format. A document read from <c>.csv</c>
/// keeps the separator and encoding it was read with, and a <c>.csv</c> save writes them back (§9.15).
/// </summary>
/// <param name="Path">The file, or null for a new document that was never saved (Save then asks for a path).</param>
/// <param name="Format">The format Save writes.</param>
/// <param name="CsvFormat">The separator and encoding of the <c>.csv</c> file the document was read from, or null when it was not read from one.</param>
public sealed record DocumentFile(string? Path, FileFormat Format, CsvFormat? CsvFormat)
{
    /// <summary>A new document: no file yet, saved as <c>.xlsx</c> unless the user chooses <c>.csv</c>.</summary>
    public static DocumentFile Untitled { get; } = new(null, FileFormat.Xlsx, null);

    /// <summary>True for a document that was never saved: Save asks for a path.</summary>
    public bool IsUntitled => Path is null;

    /// <summary>The folder of <see cref="Path"/>, or null for an untitled document.</summary>
    public string? Folder => Path is null ? null : System.IO.Path.GetDirectoryName(Path);

    /// <summary>The format a <c>.csv</c> save writes: the one the document was read with, otherwise <see cref="CsvFormat.Default"/> (§5.9).</summary>
    public CsvFormat CsvFormatForSave => CsvFormat ?? Csv.CsvFormat.Default;

    /// <summary>The name the window title shows: the file name with its extension, or <paramref name="untitledName"/>.</summary>
    /// <param name="untitledName">The app's localised name for a document that was never saved.</param>
    public string DisplayName(string untitledName) => Path is null ? untitledName : System.IO.Path.GetFileName(Path);

    /// <summary>The name the save picker suggests: the file name without its extension (the picker adds the chosen type's), or <paramref name="untitledName"/>.</summary>
    /// <param name="untitledName">The app's localised name for a document that was never saved.</param>
    public string SuggestedFileName(string untitledName) => Path is null ? untitledName : System.IO.Path.GetFileNameWithoutExtension(Path);

    /// <summary>The same document after it was saved to <paramref name="path"/> in <paramref name="format"/>; the read <see cref="CsvFormat"/> is kept.</summary>
    public DocumentFile SavedAs(string path, FileFormat format) => this with { Path = path, Format = format };
}
