# Architecture

Two projects (ADR 0002): a portable core with all the rules, and a thin WinUI 3 app that renders it (ADR 0001).

```
SpreadsheetEditor.App (WinUI 3, net10.0-windows)
  Window, title bar, menus, pickers, clipboard, drag and drop, settings, dialogs, resources
  ViewModels/  one per document, sheet, card, entry; bind with x:Bind
        │ references
        ▼
SpreadsheetEditor.Core (net10.0, no Windows API)
  Model/       Workbook, Sheet, Column, Card, Entry, CellStyle, EntryImage, ImageLayout, RgbColor, CellValue, NumberFormat
  Values/      a value's text for a culture and edited text parsed back; date serials (1900 and 1904); date and time number formats
  Xlsx/        reader (table detection, cards, styles, theme colours, pictures) and writer (in-place edit of the package)
  Csv/         reader and writer
  Imaging/     picture header facts (pixel size, format), the cell layout of pictures, crop rectangles
  Formulas/    A1 references in formula text: shifted for inserted or deleted rows, translated for shared formulas, renamed sheets
  Editing/     WorkbookEditor: operations on the model (cards and entries; values, sheets, table name, picture layout to come) with undo,
               redo, the modified flag and change events (WorkbookChange: CardInserted, CardRemoved, CardMoved, EntriesChanged,
               EntryRenamed)
```

## The document
A `Workbook` holds the sheets that have a table, in file order, and the source it was read from. A `Sheet` holds its table's position in
the worksheet (first column, header row, first data row, optional table name), its `Column`s, its `Card`s and its `ImageLayout`
(picture width and height in pixels, pictures per row in a cell). A `Card` is one table row (or a group of rows merged vertically) and
holds one `Entry` per column. An `Entry` has its text, its original typed value (number, date, boolean, text), an optional `CellStyle`
(fill, text colour, font, alignment) and its pictures (`EntryImage`: bytes, format, pixel size). `docs/spec.md` defines every rule.

## Editing
The app edits the document only through `WorkbookEditor`, on the UI thread (the core is not thread-safe, ADR 0002). Each operation is an
undoable step on one history: a new step drops the redo steps, and loading clears the history. Cards are identified by the `Card`
object, entries by their `Column` object (spec §9.36); a row span belongs to its card and moves with it, a new card has span 1 and a
duplicate keeps its source's span but, like every new card, has no source row (`Card.SourceFirstRow` is null), which is how the writer
tells new cards from read ones. Consecutive renames of one entry (the app sends one per keystroke) are a single step until another
operation, an undo or redo, a save or `EndTyping` (focus leaves the field). The modified flag compares the undo depth with the depth at
the last load or save (`MarkSaved`), so undoing back to the saved state clears it; once a new step drops the redo steps that led back to
that state, the flag stays set until the next save. After every step, undo and redo, `Changed` carries what changed (an undo raises the
inverse change) and `StateChanged` says that `CanUndo`, `CanRedo` or `IsModified` may have changed.

## Files
- `.xlsx`: the reader opens the package with the Open XML SDK and keeps it; the writer edits that same package: cell values and styles
  in the table, the rows it inserts or merges for pictures, the drawing anchors of the pictures it owns, the print area and the layout
  metadata cells. Everything else stays as read. A new workbook starts from an empty package.
- `.csv`: one sheet; text only.
- Opening and saving run on the thread pool and hand a finished `Workbook` (or an error) back to the UI thread.

## The app
- `MainWindow`: TitleBar, MenuBar, one TabView tab per sheet, an InfoBar for status. Commands have keyboard accelerators.
- A sheet's tab: a settings bar (table name, picture size and per-row count) and a virtualised list of cards (ItemsRepeater in a
  ScrollViewer, or ListView), each card a form of its entries.
- Pictures are decoded at their display size (`DecodePixelWidth`) and only when realised.
- Settings (language, window placement, recent files) are stored per user (ApplicationData for unpackaged apps or a JSON file in
  `%LOCALAPPDATA%\SpreadsheetEditor`; decided by the task that adds settings).
