namespace SpreadsheetEditor.Core.Imaging;

/// <summary>The rows a card needs for its pictures and where each picture goes (docs/spec.md §5.3, §5.4, §9.25, §9.26).</summary>
/// <param name="HasPictures">True when any value of the card has a picture.</param>
/// <param name="PictureRows">V: the most picture rows any value of the card needs, at least 1.</param>
/// <param name="PictureWorksheetRows">The worksheet rows the pictures need, ceil(V / r), whatever the card's row span.</param>
/// <param name="WorksheetRowsNeeded">
/// max(row span, <paramref name="PictureWorksheetRows"/>): the rows the card takes when rows are only inserted.
/// </param>
/// <param name="RowHeightsPoints">
/// One entry per row of <paramref name="WorksheetRowsNeeded"/>: the height to set in points, or null to leave the row's height alone
/// (a card without pictures, or a row below the last picture row).
/// </param>
/// <param name="Placements">One list per table column, in order, holding that value's pictures in order.</param>
public sealed record CardPictureLayout(
    bool HasPictures,
    int PictureRows,
    int PictureWorksheetRows,
    int WorksheetRowsNeeded,
    IReadOnlyList<double?> RowHeightsPoints,
    IReadOnlyList<IReadOnlyList<PicturePlacement>> Placements);
