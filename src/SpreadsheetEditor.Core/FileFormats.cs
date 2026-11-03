namespace SpreadsheetEditor.Core;

/// <summary>Decides a file's format from its name (docs/spec.md §9.35).</summary>
public static class FileFormats
{
    /// <summary>
    /// The format to save or read <paramref name="path"/> in: <c>.csv</c> and <c>.xlsx</c> by extension, ignoring case; a path with no
    /// extension, or any other extension, takes the format of the filter the user chose.
    /// </summary>
    public static FileFormat FromPath(string path, FileFormat chosenFilter)
    {
        ArgumentNullException.ThrowIfNull(path);
        var extension = Path.GetExtension(path);
        if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
        {
            return FileFormat.Csv;
        }
        if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return FileFormat.Xlsx;
        }
        return chosenFilter;
    }
}
