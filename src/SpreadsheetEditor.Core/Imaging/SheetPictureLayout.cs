namespace SpreadsheetEditor.Core.Imaging;

/// <summary>The picture layout of a sheet as a whole, the same for every card (docs/spec.md §5.2, §9.26).</summary>
/// <param name="Width">One picture unit's width in pixels (w).</param>
/// <param name="Height">One picture row's height in pixels (h).</param>
/// <param name="PicturesPerRow">Pictures per row after the cap at <see cref="ColumnWidths.MaxPx"/> (P); this is what the metadata cell holds.</param>
/// <param name="PictureRowsPerWorksheetRow">How many picture rows one worksheet row holds (r = max(1, floor(540 / h))).</param>
/// <param name="Columns">One entry per table column, in order.</param>
public sealed record SheetPictureLayout(
    int Width,
    int Height,
    int PicturesPerRow,
    int PictureRowsPerWorksheetRow,
    IReadOnlyList<ColumnPictureLayout> Columns);
