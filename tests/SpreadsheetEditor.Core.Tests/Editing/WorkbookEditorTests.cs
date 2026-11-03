using SpreadsheetEditor.Core.Editing;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Editing;

[TestClass]
public sealed partial class WorkbookEditorTests
{
    private static readonly string[] ExpectedTypingSteps = ["111 2", "111 2", "11 2", "1 2", "1 y", "x y"];

    // Cards.

    [TestMethod]
    public void AddCardAppendsAnEmptyNewCard()
    {
        var (editor, sheet) = Open(cards: 2);

        var card = editor.AddCard(sheet);

        Assert.AreSame(card, sheet.Cards[^1]);
        Assert.AreEqual("r0a r0b|r1a r1b|_ _", Describe(sheet));
        Assert.AreEqual(1, card.SourceRowSpan);
        Assert.IsNull(card.SourceFirstRow);
        Assert.IsTrue(card.Entries.All(e => e.Original is null && e.Style is null && e.Images.Count == 0));
        AssertUndoRedo(editor, sheet, "r0a r0b|r1a r1b", "r0a r0b|r1a r1b|_ _");
    }

    [TestMethod]
    public void InsertsCardsAboveAndBelow()
    {
        var (editor, sheet) = Open(cards: 2);

        editor.InsertCardAbove(sheet, sheet.Cards[0]);
        Assert.AreEqual("_ _|r0a r0b|r1a r1b", Describe(sheet));
        editor.InsertCardBelow(sheet, sheet.Cards[2]);
        Assert.AreEqual("_ _|r0a r0b|r1a r1b|_ _", Describe(sheet));

        AssertUndoRedo(editor, sheet, "_ _|r0a r0b|r1a r1b", "_ _|r0a r0b|r1a r1b|_ _");
        editor.Undo();
        editor.Undo();
        Assert.AreEqual("r0a r0b|r1a r1b", Describe(sheet));
    }

    [TestMethod]
    public void DuplicateCopiesValuesStylesPicturesAndSpanIntoANewCard()
    {
        var (editor, sheet) = Open(cards: 2);
        var source = sheet.Cards[0];
        source.SourceRowSpan = 3;
        source.SourceFirstRow = 2;
        var style = new CellStyle { Fill = new RgbColor(1, 2, 3), Font = new FontStyle { Bold = true } };
        var value = new CellValue(CellValueKind.Number, "r0a", Number: 5);
        var image = new EntryImage([1, 2, 3], ImageFormat.Png, 10, 20);
        source.Entries[0].Style = style;
        source.Entries[0].Original = value;
        source.Entries[0].Images.Add(image);

        var copy = editor.DuplicateCard(sheet, source);

        Assert.AreSame(copy, sheet.Cards[1]);
        Assert.AreEqual("r0a r0b|r0a r0b|r1a r1b", Describe(sheet));
        Assert.AreEqual(3, copy.SourceRowSpan);
        Assert.IsNull(copy.SourceFirstRow, "a duplicate is a new card");
        Assert.AreEqual(2, copy.CopiedFromRow, "it copies the source card's row as read");
        Assert.AreEqual(2, source.SourceFirstRow);
        Assert.AreNotSame(source.Entries[0], copy.Entries[0]);
        Assert.AreSame(style, copy.Entries[0].Style);
        Assert.AreSame(value, copy.Entries[0].Original);
        Assert.AreNotSame(source.Entries[0].Images, copy.Entries[0].Images);
        CollectionAssert.AreEqual(new[] { image }, copy.Entries[0].Images);
        AssertUndoRedo(editor, sheet, "r0a r0b|r1a r1b", "r0a r0b|r0a r0b|r1a r1b");
    }

    [TestMethod]
    public void ACopyOfACopyCopiesTheFirstCardsRow()
    {
        var (editor, sheet) = Open(cards: 2);
        sheet.Cards[0].SourceFirstRow = 4;
        var copy = editor.DuplicateCard(sheet, sheet.Cards[0]);

        var second = editor.DuplicateCard(sheet, copy);
        var fresh = editor.DuplicateCard(sheet, editor.InsertCardBelow(sheet, sheet.Cards[^1]));

        Assert.AreEqual(4, second.CopiedFromRow);
        Assert.IsNull(fresh.CopiedFromRow, "a copy of a card that was not read from the file has no row to copy from");
        editor.Undo();
        editor.Undo();
        editor.Undo();
        editor.Redo();
        Assert.AreSame(second, sheet.Cards[2]);
        Assert.AreEqual(4, sheet.Cards[2].CopiedFromRow);
    }

    [TestMethod]
    public void RemoveCardAndUndoPutsTheSameCardBack()
    {
        var (editor, sheet) = Open(cards: 3);
        var card = sheet.Cards[1];

        editor.RemoveCard(sheet, card);

        Assert.AreEqual("r0a r0b|r2a r2b", Describe(sheet));
        editor.Undo();
        Assert.AreSame(card, sheet.Cards[1]);
        editor.Redo();
        Assert.AreEqual("r0a r0b|r2a r2b", Describe(sheet));
    }

    [TestMethod]
    public void NeverRemovesTheLastCard()
    {
        var (editor, sheet) = Open(cards: 1);

        Assert.IsFalse(editor.CanRemoveCard(sheet));
        var error = Assert.ThrowsExactly<InvalidOperationException>(() => editor.RemoveCard(sheet, sheet.Cards[0]));
        Assert.AreEqual("Cannot remove the last card.", error.Message);
        Assert.AreEqual(1, sheet.Cards.Count);
        Assert.IsFalse(editor.CanUndo);
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void MovesCardsToAnIndexUpAndDown()
    {
        var (editor, sheet) = Open(cards: 4);

        Assert.IsTrue(editor.MoveCard(sheet, sheet.Cards[0], 2));
        Assert.AreEqual("r1a r1b|r2a r2b|r0a r0b|r3a r3b", Describe(sheet));
        Assert.IsTrue(editor.MoveCardUp(sheet, sheet.Cards[3]));
        Assert.AreEqual("r1a r1b|r2a r2b|r3a r3b|r0a r0b", Describe(sheet));
        Assert.IsTrue(editor.MoveCardDown(sheet, sheet.Cards[0]));
        Assert.AreEqual("r2a r2b|r1a r1b|r3a r3b|r0a r0b", Describe(sheet));

        editor.Undo();
        editor.Undo();
        editor.Undo();
        Assert.AreEqual("r0a r0b|r1a r1b|r2a r2b|r3a r3b", Describe(sheet));
        editor.Redo();
        editor.Redo();
        editor.Redo();
        Assert.AreEqual("r2a r2b|r1a r1b|r3a r3b|r0a r0b", Describe(sheet));
    }

    [TestMethod]
    public void MovesAtTheEndsDoNothingAndAreNotRecorded()
    {
        var (editor, sheet) = Open(cards: 2);

        Assert.IsFalse(editor.CanMoveCardUp(sheet, sheet.Cards[0]));
        Assert.IsTrue(editor.CanMoveCardDown(sheet, sheet.Cards[0]));
        Assert.IsTrue(editor.CanMoveCardUp(sheet, sheet.Cards[1]));
        Assert.IsFalse(editor.CanMoveCardDown(sheet, sheet.Cards[1]));
        Assert.IsFalse(editor.MoveCardUp(sheet, sheet.Cards[0]));
        Assert.IsFalse(editor.MoveCardDown(sheet, sheet.Cards[1]));
        Assert.IsFalse(editor.MoveCard(sheet, sheet.Cards[1], 1));

        Assert.AreEqual("r0a r0b|r1a r1b", Describe(sheet));
        Assert.IsFalse(editor.CanUndo);
        Assert.IsFalse(editor.IsModified);
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => editor.MoveCard(sheet, sheet.Cards[0], 2));
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => editor.MoveCard(sheet, sheet.Cards[0], -1));
    }

    [TestMethod]
    public void RowSpansMoveWithTheirCards()
    {
        var (editor, sheet) = Open(cards: 4);
        for (var i = 0; i < 4; i++)
        {
            sheet.Cards[i].SourceRowSpan = i + 1;
            sheet.Cards[i].SourceFirstRow = 10 * (i + 1);
        }

        editor.MoveCard(sheet, sheet.Cards[3], 0);
        editor.InsertCardAbove(sheet, sheet.Cards[1]);
        editor.RemoveCard(sheet, sheet.Cards[3]);
        editor.DuplicateCard(sheet, sheet.Cards[0]);
        editor.ReorderEntries(sheet, [sheet.Columns[1], sheet.Columns[0]]);
        editor.AddEntry(sheet);

        Assert.AreEqual("4@40 4@- 1@- 1@10 3@30", Spans(sheet));
        while (editor.CanUndo)
        {
            editor.Undo();
        }

        Assert.AreEqual("1@10 2@20 3@30 4@40", Spans(sheet));
        while (editor.CanRedo)
        {
            editor.Redo();
        }

        Assert.AreEqual("4@40 4@- 1@- 1@10 3@30", Spans(sheet));
    }

    [TestMethod]
    public void RejectsCardsAndSheetsFromElsewhere()
    {
        var (editor, sheet) = Open(cards: 2);
        var (_, other) = Open(cards: 2);

        Assert.ThrowsExactly<ArgumentException>(() => editor.RemoveCard(sheet, other.Cards[0]));
        Assert.ThrowsExactly<ArgumentException>(() => editor.AddCard(other));
        Assert.ThrowsExactly<ArgumentException>(() => editor.RenameEntry(sheet, other.Columns[0], "x"));
        Assert.IsFalse(editor.CanUndo);
    }

    // Entries.

    [TestMethod]
    public void AddEntryAppendsAnUnnamedColumnAndAnEmptyEntryToEveryCard()
    {
        var (editor, sheet) = Open(cards: 2);

        var column = editor.AddEntry(sheet);

        Assert.AreSame(column, sheet.Columns[^1]);
        Assert.AreEqual("", column.Name);
        Assert.AreEqual("a b ", Names(sheet));
        Assert.AreEqual("r0a r0b _|r1a r1b _", Describe(sheet));
        AssertUndoRedo(editor, sheet, "r0a r0b|r1a r1b", "r0a r0b _|r1a r1b _");
        Assert.AreSame(column, sheet.Columns[^1]);
    }

    [TestMethod]
    public void RemovesEntriesByIdentityWithDuplicateAndEmptyNames()
    {
        var (editor, sheet) = Open(cards: 2, "x", "", "x", "");

        editor.RemoveEntries(sheet, [sheet.Columns[2], sheet.Columns[1]]);

        Assert.AreEqual("r0a r0d|r1a r1d", Describe(sheet));
        Assert.AreEqual("x ", Names(sheet));
        editor.Undo();
        Assert.AreEqual("r0a r0b r0c r0d|r1a r1b r1c r1d", Describe(sheet));
        Assert.AreEqual("x  x ", Names(sheet));
        editor.Redo();
        Assert.AreEqual("r0a r0d|r1a r1d", Describe(sheet));
    }

    [TestMethod]
    public void UndoOfRemoveEntriesPutsTheSameColumnsBack()
    {
        var (editor, sheet) = Open(cards: 1, "x", "x", "x");
        var columns = sheet.Columns.ToArray();
        var entries = sheet.Cards[0].Entries.ToArray();

        editor.RemoveEntries(sheet, [columns[0], columns[2]]);
        Assert.AreSame(columns[1], sheet.Columns.Single());
        editor.Undo();

        CollectionAssert.AreEqual(columns, sheet.Columns);
        CollectionAssert.AreEqual(entries, sheet.Cards[0].Entries);
    }

    [TestMethod]
    public void NeverRemovesTheLastEntry()
    {
        var (editor, sheet) = Open(cards: 2);
        var all = sheet.Columns.ToArray();

        Assert.IsFalse(editor.CanRemoveEntries(sheet, all));
        Assert.IsFalse(editor.CanRemoveEntries(sheet, []));
        Assert.IsTrue(editor.CanRemoveEntries(sheet, [all[0]]));
        Assert.ThrowsExactly<InvalidOperationException>(() => editor.RemoveEntries(sheet, all));
        Assert.ThrowsExactly<ArgumentException>(() => editor.RemoveEntries(sheet, []));

        editor.RemoveEntries(sheet, [all[0]]);
        Assert.IsFalse(editor.CanRemoveEntries(sheet, [all[1]]));
        Assert.ThrowsExactly<InvalidOperationException>(() => editor.RemoveEntries(sheet, [all[1]]));
        Assert.AreEqual("r0b|r1b", Describe(sheet));
    }

    [TestMethod]
    public void ReordersEntriesByIdentityInEveryCard()
    {
        var (editor, sheet) = Open(cards: 2, "x", "", "x");
        var c = sheet.Columns.ToArray();

        Assert.IsTrue(editor.ReorderEntries(sheet, [c[2], c[0], c[1]]));

        CollectionAssert.AreEqual(new[] { c[2], c[0], c[1] }, sheet.Columns);
        Assert.AreEqual("r0c r0a r0b|r1c r1a r1b", Describe(sheet));
        AssertUndoRedo(editor, sheet, "r0a r0b r0c|r1a r1b r1c", "r0c r0a r0b|r1c r1a r1b");
    }

    [TestMethod]
    public void ReorderRejectsAnythingButAPermutation()
    {
        var (editor, sheet) = Open(cards: 1, "x", "y", "z");
        var c = sheet.Columns.ToArray();

        Assert.ThrowsExactly<ArgumentException>(() => editor.ReorderEntries(sheet, [c[0], c[1]]));
        Assert.ThrowsExactly<ArgumentException>(() => editor.ReorderEntries(sheet, [c[0], c[0], c[1]]));
        Assert.ThrowsExactly<ArgumentException>(() => editor.ReorderEntries(sheet, [c[0], c[1], new Column("z")]));
        Assert.IsFalse(editor.ReorderEntries(sheet, c));
        Assert.IsFalse(editor.CanUndo);
    }

    [TestMethod]
    public void RenamesAnEntryByIdentity()
    {
        var (editor, sheet) = Open(cards: 1, "x", "x");

        Assert.IsTrue(editor.RenameEntry(sheet, sheet.Columns[1], "Unnamed: 0"));

        Assert.AreEqual("x Unnamed: 0", Names(sheet));
        Assert.IsTrue(editor.RenameEntry(sheet, sheet.Columns[0], "Unnamed: 0"));
        Assert.AreEqual("Unnamed: 0 Unnamed: 0", Names(sheet), "names may repeat; only the empty name is unnamed");
        Assert.IsFalse(editor.RenameEntry(sheet, sheet.Columns[0], "Unnamed: 0"));
        editor.Undo();
        Assert.AreEqual("x Unnamed: 0", Names(sheet));
        editor.Undo();
        Assert.AreEqual("x x", Names(sheet));
        editor.Redo();
        Assert.AreEqual("x Unnamed: 0", Names(sheet));
    }

    [TestMethod]
    public void TypingANameIsOneUndoStep()
    {
        var (editor, sheet) = Open(cards: 1, "x", "y");
        var column = sheet.Columns[0];

        editor.RenameEntry(sheet, column, "N");
        editor.RenameEntry(sheet, column, "Na");
        editor.RenameEntry(sheet, column, "Nam");
        editor.Undo();

        Assert.AreEqual("x", column.Name);
        Assert.IsFalse(editor.CanUndo);
        editor.Redo();
        Assert.AreEqual("Nam", column.Name);
    }

    [TestMethod]
    public void TypingStepsBreakOnAnotherEntryAnotherOperationAndEndTyping()
    {
        var (editor, sheet) = Open(cards: 1, "x", "y");
        var (first, second) = (sheet.Columns[0], sheet.Columns[1]);

        editor.RenameEntry(sheet, first, "1");
        editor.RenameEntry(sheet, second, "2");
        editor.RenameEntry(sheet, first, "11");
        editor.EndTyping();
        editor.RenameEntry(sheet, first, "111");
        editor.AddCard(sheet);
        editor.RenameEntry(sheet, first, "1111");

        var names = new List<string>();
        while (editor.CanUndo)
        {
            editor.Undo();
            names.Add(Names(sheet));
        }

        CollectionAssert.AreEqual(ExpectedTypingSteps, names);
    }

    [TestMethod]
    public void HiddenHeaderEntriesCannotBeRenamedAndNewOnesAreUnnamed()
    {
        var (editor, sheet) = Open(cards: 1);
        sheet.HeaderRow = null;

        Assert.ThrowsExactly<InvalidOperationException>(() => editor.RenameEntry(sheet, sheet.Columns[0], "x"));
        var column = editor.AddEntry(sheet);

        Assert.AreEqual("", column.Name);
        Assert.AreEqual("a b ", Names(sheet));
    }

    // Undo, redo and the modified flag.

    [TestMethod]
    public void ANewOperationDropsRedo()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.AddCard(sheet);
        editor.AddEntry(sheet);
        editor.Undo();
        Assert.IsTrue(editor.CanRedo);

        editor.AddCard(sheet);

        Assert.IsFalse(editor.CanRedo);
        Assert.AreEqual("r0a r0b|_ _|_ _", Describe(sheet));
        Assert.ThrowsExactly<InvalidOperationException>(editor.Redo);
        editor.Undo();
        editor.Undo();
        Assert.ThrowsExactly<InvalidOperationException>(editor.Undo);
        Assert.AreEqual("r0a r0b", Describe(sheet));
    }

    [TestMethod]
    public void EveryOperationSetsTheModifiedFlag()
    {
        var operations = new (string Name, Action<WorkbookEditor, Sheet> Do)[]
        {
            ("add card", (e, s) => e.AddCard(s)),
            ("insert above", (e, s) => e.InsertCardAbove(s, s.Cards[0])),
            ("insert below", (e, s) => e.InsertCardBelow(s, s.Cards[0])),
            ("duplicate", (e, s) => e.DuplicateCard(s, s.Cards[0])),
            ("remove card", (e, s) => e.RemoveCard(s, s.Cards[0])),
            ("move", (e, s) => e.MoveCard(s, s.Cards[0], 1)),
            ("move up", (e, s) => e.MoveCardUp(s, s.Cards[1])),
            ("move down", (e, s) => e.MoveCardDown(s, s.Cards[0])),
            ("add entry", (e, s) => e.AddEntry(s)),
            ("remove entry", (e, s) => e.RemoveEntries(s, [s.Columns[0]])),
            ("reorder", (e, s) => e.ReorderEntries(s, [s.Columns[1], s.Columns[0]])),
            ("rename", (e, s) => e.RenameEntry(s, s.Columns[0], "z")),
        };

        foreach (var (name, operation) in operations)
        {
            var (editor, sheet) = Open(cards: 2);
            var before = Describe(sheet) + "/" + Names(sheet);

            operation(editor, sheet);
            Assert.IsTrue(editor.IsModified, name);
            Assert.IsTrue(editor.CanUndo, name);

            editor.Undo();
            Assert.IsFalse(editor.IsModified, name);
            Assert.AreEqual(before, Describe(sheet) + "/" + Names(sheet), name);
        }
    }

    [TestMethod]
    public void SavingClearsTheModifiedFlagAndUndoToTheSavedStateClearsItAgain()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.AddCard(sheet);
        editor.MarkSaved();
        Assert.IsFalse(editor.IsModified);
        Assert.IsTrue(editor.CanUndo, "saving keeps the history");

        editor.AddCard(sheet);
        Assert.IsTrue(editor.IsModified);
        editor.Undo();
        Assert.IsFalse(editor.IsModified);
        editor.Undo();
        Assert.IsTrue(editor.IsModified);
        editor.Redo();
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void TheSavedStateDroppedWithRedoStaysModified()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.AddCard(sheet);
        editor.MarkSaved();
        editor.Undo();

        editor.AddEntry(sheet);
        editor.Undo();

        Assert.IsTrue(editor.IsModified);
        Assert.IsFalse(editor.CanUndo);
    }

    [TestMethod]
    public void TypingAfterASaveIsANewStep()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.RenameEntry(sheet, sheet.Columns[0], "N");
        editor.MarkSaved();
        editor.RenameEntry(sheet, sheet.Columns[0], "Na");

        editor.Undo();

        Assert.AreEqual("N", sheet.Columns[0].Name);
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void LoadClearsTheHistoryAndTheModifiedFlag()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.AddCard(sheet);
        editor.AddCard(sheet);
        editor.Undo();
        var (_, newSheet) = Open(cards: 1);
        var workbook = new Workbook();
        workbook.Sheets.Add(newSheet);
        var raised = 0;
        editor.StateChanged += (_, _) => raised++;

        editor.Load(workbook);

        Assert.AreSame(workbook, editor.Workbook);
        Assert.IsFalse(editor.CanUndo);
        Assert.IsFalse(editor.CanRedo);
        Assert.IsFalse(editor.IsModified);
        Assert.AreEqual(1, raised);
    }

    // Change notifications.

    [TestMethod]
    public void RaisesChangesForDoUndoAndRedo()
    {
        var (editor, sheet) = Open(cards: 3);
        var changes = new List<WorkbookChange>();
        var states = 0;
        editor.Changed += (sender, e) =>
        {
            Assert.AreSame(editor, sender);
            changes.AddRange(e.Changes);
        };
        editor.StateChanged += (_, _) => states++;
        var card = sheet.Cards[0];

        editor.MoveCard(sheet, card, 2);
        editor.Undo();
        editor.Redo();
        var added = editor.AddCard(sheet);
        editor.Undo();
        editor.RemoveCard(sheet, card);
        editor.Undo();
        var column = editor.AddEntry(sheet);
        editor.RenameEntry(sheet, column, "n");

        CollectionAssert.AreEqual(
            new WorkbookChange[]
            {
                new CardMoved(sheet, card, 0, 2),
                new CardMoved(sheet, card, 2, 0),
                new CardMoved(sheet, card, 0, 2),
                new CardInserted(sheet, 3, added),
                new CardRemoved(sheet, 3, added),
                new CardRemoved(sheet, 2, card),
                new CardInserted(sheet, 2, card),
                new EntriesChanged(sheet),
                new EntryRenamed(sheet, column),
            },
            changes);
        Assert.AreEqual(9, states);
    }

    [TestMethod]
    public void TheModelHasChangedWhenChangesAreRaised()
    {
        var (editor, sheet) = Open(cards: 1);
        var count = 0;
        editor.Changed += (_, _) => count = sheet.Cards.Count;

        editor.AddCard(sheet);

        Assert.AreEqual(2, count);
    }

    // Modified flag.

    [TestMethod]
    public void MarkModifiedMarksAnUneditedDocumentWithoutAnUndoStep()
    {
        var (editor, _) = Open(cards: 1);
        var raised = 0;
        editor.StateChanged += (_, _) => raised++;

        editor.MarkModified();

        Assert.IsTrue(editor.IsModified);
        Assert.IsFalse(editor.CanUndo);
        Assert.AreEqual(1, raised);
    }

    [TestMethod]
    public void MarkModifiedIsClearedBySaveAndLoad()
    {
        var (editor, _) = Open(cards: 1);

        editor.MarkModified();
        editor.MarkSaved();
        Assert.IsFalse(editor.IsModified);

        editor.MarkModified();
        editor.Load(new Workbook());
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void MarkModifiedHoldsThroughAnUndoRedoRoundTrip()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.AddCard(sheet);
        editor.MarkSaved();
        editor.MarkModified();

        editor.Undo();
        editor.Redo();

        Assert.IsTrue(editor.IsModified);
        editor.MarkSaved();
        Assert.IsFalse(editor.IsModified);
    }

    // Helpers.

    /// <summary>An editor over a workbook with one sheet: the given column names (default a, b), each card's entries "r{card}{column letter}".</summary>
    private static (WorkbookEditor Editor, Sheet Sheet) Open(int cards, params string[] names)
    {
        if (names.Length == 0)
        {
            names = ["a", "b"];
        }

        var sheet = new Sheet("S");
        foreach (var name in names)
        {
            sheet.Columns.Add(new Column(name));
        }

        for (var r = 0; r < cards; r++)
        {
            var card = new Card();
            for (var c = 0; c < names.Length; c++)
            {
                card.Entries.Add(new Entry { Text = $"r{r}{(char)('a' + c)}" });
            }

            sheet.Cards.Add(card);
        }

        var workbook = new Workbook();
        workbook.Sheets.Add(sheet);
        return (new WorkbookEditor(workbook), sheet);
    }

    /// <summary>Each card's entry texts ("_" for empty), cards joined by "|".</summary>
    private static string Describe(Sheet sheet)
    {
        foreach (var card in sheet.Cards)
        {
            Assert.AreEqual(sheet.Columns.Count, card.Entries.Count, "every card has one entry per column");
        }

        return string.Join("|", sheet.Cards.Select(card => string.Join(" ", card.Entries.Select(e => e.Text.Length == 0 ? "_" : e.Text))));
    }

    private static string Names(Sheet sheet) => string.Join(" ", sheet.Columns.Select(c => c.Name));

    private static string Spans(Sheet sheet) =>
        string.Join(" ", sheet.Cards.Select(c => $"{c.SourceRowSpan}@{(c.SourceFirstRow is { } row ? row.ToString(System.Globalization.CultureInfo.InvariantCulture) : "-")}"));

    /// <summary>Undoes the last step (expecting <paramref name="before"/>) and redoes it (expecting <paramref name="after"/>).</summary>
    private static void AssertUndoRedo(WorkbookEditor editor, Sheet sheet, string before, string after)
    {
        editor.Undo();
        Assert.AreEqual(before, Describe(sheet));
        Assert.IsTrue(editor.CanRedo);
        editor.Redo();
        Assert.AreEqual(after, Describe(sheet));
        Assert.IsFalse(editor.CanRedo);
    }
}
