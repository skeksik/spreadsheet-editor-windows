using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>Renames a column. Consecutive renames of the same column merge into one step that keeps the first old name.</summary>
internal sealed class RenameEntryEdit : IEdit
{
    private readonly Sheet sheet;
    private readonly Column column;
    private readonly string oldName;
    private string newName;

    public RenameEntryEdit(Sheet sheet, Column column, string oldName, string newName)
    {
        this.sheet = sheet;
        this.column = column;
        this.oldName = oldName;
        this.newName = newName;
    }

    public IReadOnlyList<WorkbookChange> Apply()
    {
        column.Name = newName;
        return [new EntryRenamed(sheet, column)];
    }

    public IReadOnlyList<WorkbookChange> Revert()
    {
        column.Name = oldName;
        return [new EntryRenamed(sheet, column)];
    }

    public bool TryMerge(IEdit next)
    {
        if (next is not RenameEntryEdit rename || rename.column != column)
        {
            return false;
        }

        newName = rename.newName;
        return true;
    }
}
