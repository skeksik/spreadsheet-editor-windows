# Architecture

Two projects (ADR 0002): a portable core with all the rules, and a thin WinUI 3 app that renders it (ADR 0001).

```
SpreadsheetEditor.App (WinUI 3, net10.0-windows)
  Window, title bar, menus, pickers, clipboard, drag and drop, settings, dialogs, resources
  ViewModels/  one per document, sheet, card, entry; bind with x:Bind
        │ references
        ▼
SpreadsheetEditor.Core (net10.0, no Windows API)
  Model/       Workbook, Sheet, Column, Card, Entry, CellStyle, EntryImage, ImageLayout, RgbColor
  Xlsx/        reader (table detection, cards, styles, theme colours, pictures) and writer (in-place edit of the package)
  Csv/         reader and writer
  Imaging/     picture header facts (pixel size, format), the cell layout of pictures, crop rectangles
  Editing/     operations on the model (add, remove, move, duplicate cards and entries; edit text, style, pictures) with undo and redo
```

## The document
A `Workbook` holds the sheets that have a table, in file order, and the source it was read from. A `Sheet` holds its table's position in
the worksheet (first column, header row, first data row, optional table name), its `Column`s, its `Card`s and its `ImageLayout`
(picture width and height in pixels, pictures per row in a cell). A `Card` is one table row (or a group of rows merged vertically) and
holds one `Entry` per column. An `Entry` has its text, its original typed value (number, date, boolean, text), an optional `CellStyle`
(fill, text colour, font, alignment) and its pictures (`EntryImage`: bytes, format, pixel size). `docs/spec.md` defines every rule.

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
