# Roadmap

Each unchecked item is one task, one branch and one PR unless it says otherwise. "Cloud" tasks are built by cloud workers, with CI
building what a container cannot. "Local" tasks need the owner's PC (running the app, looking at it, the owner's own files). Spec
references are to `docs/spec.md`.

## M0: Foundation
- [x] Repository, solution, core and app projects, tests, CI (core on Linux and Windows, app publish and smoke run), CLAUDE.md,
      ADRs 0001–0003, spec, the `/manage` command, the app icon.
- [x] Core document model (`Model/`): workbook, sheet, column, card, entry, cell value and style, pictures, picture layout.
- [x] **Owner's answers to the open questions** in spec §9 that block M1 (2, 3, 6, 7, 9, 18, 20, 26), recorded in the spec
      (2026-11-03).

## M1: Files in the core (cloud)
- [ ] `feat/picture-facts`: read a picture's format and pixel size from its header (PNG, JPEG by its SOF marker, GIF, BMP);
      `IsWide`; the centre-crop source rectangle in 1/100000 units (spec §5.4). Core only, with tests.
- [ ] `feat/picture-layout`: the cell layout of a value's pictures: units, per-column capacity, the 1762 px cap, the placement loop
      that also counts picture rows, picture rows per worksheet row from the height (spec §5.3–5.4, §9 items 25–26), worksheet
      rows needed, row heights, column widths and the px↔width formulas. Pure functions over the model (`EntryImage.IsWide`), with
      tests.
- [ ] `feat/cell-values`: a value's text for a culture and edited text parsed back: numbers, booleans, dates and times in the short
      date format, formulas as `=source`, the cell's number format kept (§4.2, §9.2–9.4). `CellValue` gains a formula and a number
      format. Core only, with tests in en-US and ru-RU.
- [ ] `feat/xlsx-read`: open an `.xlsx` with the Open XML SDK and build the model: table detection with metadata cells ignored
      and the hidden-header test (§4.1, §9.1, §9.9), values with their types, dates and formulas (§4.2, §9.2–9.3), merged-row
      grouping by table merges only (§4.3, §9.7), theme and indexed colours with tint and automatic text colour (§4.4, §9.5–9.6),
      fonts and alignment, pictures anchored to cells, layout metadata clamped (§4.5, §9.50); unsupported sheets skipped and kept
      (§9.10, §9.12), hidden worksheets read (§9.12). Keeps the opened package for the writer. Uses `feat/cell-values` for a value's
      text. Tests build workbooks in code. Two PRs: `feat/xlsx-read` (table, values, grouping, metadata, skipped sheets), then
      `feat/xlsx-read-styles` (colours, fonts, alignment, pictures).
- [ ] `feat/xlsx-write`: write the model back into the kept package in place (§5): values typed and dates kept (§9.2–9.3), styles
      with automatic text colour (§9.6), clearing removed content (§9.16–9.19), row insertion and deletion that shift what lies
      below (§9.20), a name row inserted for a new table name (§9.18), later-row values of multi-row cards kept (§9.7, §9.24),
      merges, row heights, column widths, pictures with source-rectangle crop, print area (§9.27), metadata column; a new workbook
      from an empty package; repeatable saves (§9.21). Round-trip tests. Depends on `feat/xlsx-read`, `feat/cell-values` and
      `feat/picture-layout`.
- [ ] `feat/csv-io`: read and write `.csv` (§4.7, §5.9, §9.4, §9.14–9.15, §9.34–9.35): separator and encoding detected and kept,
      one empty card for a header-only file, export of one given sheet, round-trip tests.
- [ ] `feat/sheet-names`: worksheet name rules and unique names for new and duplicated sheets (§9.39), with tests.
- [ ] `spike/jpeg-orientation` (local): does Excel apply a JPEG's EXIF orientation to a picture in an `.xlsx`? Check an
      orientation-6 JPEG on the owner's PC; if it does, `feat/picture-facts` swaps width and height for orientations 5–8. Notes only.

## M2: Editing in the core (cloud)
- [ ] `feat/editing-ops`: operations on the model with undo and redo: cards (add above/below/end, duplicate, remove, move), entries
      (add, remove by identity, reorder, rename; never the last one, §9.33, §9.36–9.38), values (text, fill, text colour incl.
      automatic, pictures add/remove/reorder), sheets (add, rename in place, duplicate, remove, reorder, §9.30–9.32), table name,
      picture layout. Row spans belong to cards (§9.22). Modified flag (§8.1). Tests.

## M3: The app (WinUI; cloud workers, CI builds; each ends with a local check card)
- [ ] `feat/app-document`: New, Open, Save, Save as with the system pickers and Ctrl+N/O/S/Shift+S; title `{file} - Spreadsheet
      Editor` with an unsaved marker; Save / Don't save / Cancel on close, open and drop; open and save off the UI thread with a
      progress state; errors in a ContentDialog with details in an expander; drop a file on the window to open it (spec §2.1–2.2,
      §10.4–10.8, §10.17). Saves write a temporary file and replace the target, so a failed save never damages the file.
- [ ] `feat/app-sheets`: one TabView tab per sheet with add, close (asks only for non-empty sheets), drag to reorder, inline rename,
      duplicate, context menu with icons and accelerators, Ctrl+T/W/Tab (§2.3, §10.1, §10.22–10.23).
- [ ] `feat/app-cards`: the sheet page: the settings bar (table name, picture width and height NumberBoxes with headers, aspect lock
      with the link glyph, pictures per row) and a virtualised card list; each card a form of entry names and multi-line values;
      add, duplicate, remove, move cards with buttons, drag and Alt+Up/Down; Enter and Tab per the owner's §10.12 answer (§1.4–1.5,
      §2.4, §2.7–2.9, §10.13, §10.19–10.20).
- [ ] `feat/app-entries`: add, remove, reorder and rename entries with ContentDialogs that work by identity (§2.5, §10.3).
- [ ] `feat/app-pictures`: the picture strip: thumbnails decoded at display size off the UI thread and cached, add from the picker,
      paste and drag and drop, preview with zoom and arrow keys, copy one or all, remove with Delete, keyboard access (§6, §9.44–9.49,
      §10.16, §10.21).
- [ ] `feat/app-styles`: fill and text colour flyouts with the system ColorPicker, recent colours, No fill and Automatic; the text
      field previews fill, colour, bold, italic and alignment if the owner agrees (§2.6, §9.43, §10.11).
- [ ] `feat/app-undo`: Edit menu and Ctrl+Z / Ctrl+Y over the core's undo; destructive actions without confirmation where undo covers
      them (§10.14).
- [ ] `feat/app-localisation`: every string in `Strings/en-US` and `Strings/ru-RU` resources (spec §7.2); follow the Windows display
      language with an override in Settings (§10.9–10.10). Overlaps every app task: run it early, then each task adds its strings.
- [ ] `feat/app-settings`: a Settings page (language, theme follows Windows), window size and position remembered, recent files in the
      File menu and the taskbar jump list (§7.3).

## M4: Release
- [ ] `spike/native-aot`: publish the app with NativeAOT and trimming; measure start time and memory against the default build; ADR.
- [ ] `test/performance`: a generated workbook with 2,000 cards and 500 pictures; open, scroll and save times recorded in
      `docs/notes/performance.md`; fix what is slow.
- [ ] `build/installer`: an installer or MSIX (ADR), the licence texts in `licences\`, a release workflow that attaches the build to a
      GitHub release. Needs the owner's choice of installer technology.
- [ ] Milestone accessibility pass with the owner: Narrator, high contrast, keyboard only (local).
