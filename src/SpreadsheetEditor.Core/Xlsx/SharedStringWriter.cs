using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Adds strings to a workbook's shared string table for the writer, reusing an existing plain item with the same text. Existing items are
/// never changed or removed, so every cell the writer does not touch keeps pointing at its string. The part is created when the workbook
/// has none, and loaded only when a string is written.
/// </summary>
internal sealed class SharedStringWriter
{
    private readonly WorkbookPart workbookPart;
    private X.SharedStringTable? table;
    private Dictionary<string, int>? indexes;
    private int count;
    private int added;

    public SharedStringWriter(WorkbookPart workbookPart) => this.workbookPart = workbookPart;

    /// <summary>The index of a shared string item holding exactly this text.</summary>
    public int Add(string text)
    {
        if (table is null || indexes is null)
        {
            Load();
        }

        added++;
        if (indexes!.TryGetValue(text, out var index))
        {
            return index;
        }

        var item = new X.SharedStringItem(new X.Text(text) { Space = SpaceProcessingModeValues.Preserve });
        table!.AppendChild(item);
        index = count++;
        indexes[text] = index;
        return index;
    }

    /// <summary>Updates the table's counts after strings were added.</summary>
    public void Finish()
    {
        if (table is null || added == 0)
        {
            return;
        }

        table.UniqueCount = (uint)count;
        if (table.Count?.Value is { } total)
        {
            // The count of references is a hint; cleared cells are not subtracted (not verified that a spreadsheet program checks it).
            table.Count = total + (uint)added;
        }
    }

    private void Load()
    {
        var part = workbookPart.SharedStringTablePart ?? workbookPart.AddNewPart<SharedStringTablePart>(RelationshipIds.Next(workbookPart));
        table = part.SharedStringTable ??= new X.SharedStringTable();
        indexes = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var item in table.Elements<X.SharedStringItem>())
        {
            // Only an item that is a single plain text can be reused: rich text and phonetic runs carry more than the text.
            if (item.ChildElements.Count == 1 && item.FirstChild is X.Text text)
            {
                indexes.TryAdd(text.Text, count);
            }

            count++;
        }
    }
}
