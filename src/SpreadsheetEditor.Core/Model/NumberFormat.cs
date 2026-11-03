namespace SpreadsheetEditor.Core.Model;

/// <summary>A cell's number format: the <c>numFmtId</c> it refers to and, for a format defined in the file, its format code.</summary>
/// <param name="Id">The number format id. Ids below 164 are built in; their code is implied rather than stored (ECMA-376 §18.8.30).</param>
/// <param name="Code">The format code as stored in the file, or null for a built-in id without one.</param>
public sealed record NumberFormat(int Id, string? Code = null)
{
    /// <summary>The <c>General</c> format, id 0.</summary>
    public static NumberFormat General { get; } = new(0);
}
