using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Which of a sheet's values keep their pictures' anchors as read and which are laid out again on save (docs/spec.md §5.2–5.4): a value
/// keeps its anchors (only moved with its rows) when its card keeps its rows, its column is in place, the sheet's picture settings are
/// those read and its pictures are those read, by their bytes and in order. Every other value's pictures, and its cell's old pictures,
/// are laid out again from <see cref="PictureLayout"/>.
/// </summary>
internal sealed class SheetPictures
{
    private readonly Sheet sheet;
    private readonly bool layoutUnchanged;
    private readonly int sourceFirstColumn;
    private readonly Dictionary<(int Row, int Column), List<EntryImage>> read = [];
    private readonly Dictionary<Card, CardPictureLayout> cards = [];
    private readonly Dictionary<(Card Card, int? Row), bool[]> kept = [];

    private SheetPictures(Sheet sheet, bool layoutUnchanged, int sourceFirstColumn)
    {
        this.sheet = sheet;
        this.layoutUnchanged = layoutUnchanged;
        this.sourceFirstColumn = sourceFirstColumn;
        Layout = PictureLayout.ForSheet(sheet);
    }

    /// <summary>The sheet-wide layout: pictures per row, picture rows per worksheet row, each column's capacity and width.</summary>
    public SheetPictureLayout Layout { get; }

    /// <summary>A sheet without a source: every value is laid out.</summary>
    public static SheetPictures ForNewSheet(Sheet sheet) => new(sheet, layoutUnchanged: false, sourceFirstColumn: 0);

    /// <summary>
    /// A sheet read from <c>.xlsx</c>: its pictures as read are the ones <see cref="SheetBuilder"/> gave its cards, by the card slots the
    /// writer found (<paramref name="slots"/>) and the table's columns as read.
    /// </summary>
    public static SheetPictures ForSource(Sheet sheet, SheetSource origin, SourceWorksheet source, List<(int FirstRow, int RowSpan)> slots)
    {
        var grid = source.Grid;
        var asRead = LayoutMetadata.Read(
            grid.Get(1, CellReference.MaxColumn)?.PlainText,
            grid.Get(2, CellReference.MaxColumn)?.PlainText,
            grid.Get(3, CellReference.MaxColumn)?.PlainText);
        var result = new SheetPictures(sheet, sheet.ImageLayout == asRead, origin.FirstColumn);
        var firstRows = slots.Select(s => s.FirstRow).ToArray();
        // The same rule as SheetBuilder.AssignPictures: by anchor row, then file order (OrderBy is stable).
        foreach (var picture in source.Pictures.OrderBy(p => p.Row))
        {
            if (picture.Column < origin.FirstColumn || picture.Column > origin.LastColumn)
            {
                continue;
            }

            var index = Array.BinarySearch(firstRows, picture.Row);
            index = index < 0 ? ~index - 1 : index;
            if (index < 0 || picture.Row >= slots[index].FirstRow + slots[index].RowSpan)
            {
                continue;
            }

            var key = (slots[index].FirstRow, picture.Column);
            if (!result.read.TryGetValue(key, out var list))
            {
                list = [];
                result.read[key] = list;
            }

            list.Add(picture.Image);
        }

        return result;
    }

    /// <summary>The card's picture layout: where each picture goes, the rows it needs and their heights.</summary>
    public CardPictureLayout ForCard(Card card)
    {
        if (!cards.TryGetValue(card, out var layout))
        {
            layout = PictureLayout.ForCard(Layout, card);
            cards[card] = layout;
        }

        return layout;
    }

    /// <summary>
    /// True when the value in table column <paramref name="column"/> keeps its pictures' anchors as read. <paramref name="sourceFirstRow"/>
    /// is the first row as read of the rows the card keeps, or null for a card in new rows.
    /// </summary>
    public bool Keeps(Card card, int? sourceFirstRow, int column)
    {
        // Comparing the bytes of every picture is not free: each value is compared once per save.
        if (!kept.TryGetValue((card, sourceFirstRow), out var values))
        {
            values = new bool[card.Entries.Count];
            for (var j = 0; j < values.Length; j++)
            {
                values[j] = Compare(card, sourceFirstRow, j);
            }

            kept[(card, sourceFirstRow)] = values;
        }

        return values[column];
    }

    private bool Compare(Card card, int? sourceFirstRow, int column)
    {
        if (!layoutUnchanged || sourceFirstRow is not { } first || sheet.Columns[column].SourceIndex != column)
        {
            return false;
        }

        var images = card.Entries[column].Images;
        var asRead = read.GetValueOrDefault((first, sourceFirstColumn + column));
        if (images.Count != (asRead?.Count ?? 0))
        {
            return false;
        }

        for (var i = 0; i < images.Count; i++)
        {
            if (!ReferenceEquals(images[i].Data, asRead![i].Data) && !images[i].Data.AsSpan().SequenceEqual(asRead[i].Data))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>True when any value of the card that is laid out again has pictures: its rows and heights then follow its layout.</summary>
    public bool LaysOutPictures(Card card, int? sourceFirstRow)
    {
        for (var j = 0; j < card.Entries.Count; j++)
        {
            if (card.Entries[j].Images.Count > 0 && !Keeps(card, sourceFirstRow, j))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>The worksheet rows the card needs for its pictures, or 0 when none of its pictures is laid out again (§5.3 step 2).</summary>
    public int RowsNeeded(Card card, int? sourceFirstRow) =>
        LaysOutPictures(card, sourceFirstRow) ? ForCard(card).WorksheetRowsNeeded : 0;
}
