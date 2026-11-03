using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Imaging;

/// <summary>Made-up pictures and sheets for the layout tests: only the format and pixel size matter, so the bytes are empty.</summary>
internal static class TestImages
{
    public static EntryImage Tall() => new([], ImageFormat.Png, 30, 40);

    public static EntryImage Wide() => new([], ImageFormat.Jpeg, 40, 30);

    public static EntryImage Square() => new([], ImageFormat.Gif, 30, 30);

    public static EntryImage UnknownFormat() => new([], ImageFormat.Unknown, 40, 30);

    public static EntryImage NoSize() => new([], ImageFormat.Png, 0, 0);

    /// <summary>A sheet with the given number of columns and cards; each card is a list of pictures per column.</summary>
    public static Sheet MakeSheet(ImageLayout layout, int columns, EntryImage[][][] cards)
    {
        var sheet = new Sheet("Sheet1") { ImageLayout = layout };
        for (var c = 0; c < columns; c++)
        {
            sheet.Columns.Add(new Column(""));
        }

        foreach (var pictures in cards)
        {
            sheet.Cards.Add(MakeCard(pictures));
        }

        return sheet;
    }

    public static Card MakeCard(EntryImage[][] pictures)
    {
        var card = new Card();
        foreach (var column in pictures)
        {
            var entry = new Entry();
            entry.Images.AddRange(column);
            card.Entries.Add(entry);
        }

        return card;
    }
}
