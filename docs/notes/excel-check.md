# Excel check

Files saved by the xlsx writer (PRs #19, #21, #22) have not been opened in Excel yet, and CI cannot do it. This check opens nine sample
workbooks, and one control file, in Excel on the owner's PC. Each file is made the way the app makes one: a source workbook is built in code, read with
`XlsxReader`, edited with `WorkbookEditor` and saved with `XlsxWriter`. Roadmap items: `test/xlsx-in-excel` and
`spike/jpeg-orientation` (M1).

CI writes the same files on every run. It also checks each one against the Open XML schemas (the Open XML SDK's validator, format
version `Microsoft365`), reads it back and checks what this list asks you to look at. So what is left to check here is what only Excel
shows: whether it asks to repair the file, the results of formulas after Excel recalculates them, and how the pictures look.

## Making the files

On Windows, with the .NET 10 SDK, in the repository folder (PowerShell):

```powershell
$env:EXCEL_CHECK_DIR = "$env:TEMP\excel-check"
dotnet test tests/SpreadsheetEditor.Core.Tests --filter "FullyQualifiedName~ExcelCheckTests"
Remove-Item Env:EXCEL_CHECK_DIR
explorer "$env:TEMP\excel-check"
```

The tests must all pass (`Passed!`). The folder then holds `00-control.xlsx` and `01-values.xlsx` … `09-jpeg-orientation.xlsx`, a few KB each. Running
the command again overwrites them. Without `EXCEL_CHECK_DIR`, the tests write the files to a temporary folder and delete them.

The code is in `tests/SpreadsheetEditor.Core.Tests/ExcelCheck/`: `ExcelCheckSamples.cs` builds the samples and `CheckPictures.cs` draws
the pictures.

## How to check each file

For every file:
1. Open it in Excel by double-clicking it. **Excel must not report a problem with the file's content or offer to repair it.** If it
   does, note the message word for word, and the repair log Excel offers (the list of what it removed or repaired).
2. Look at the sheets and cells listed below. The writer sets `fullCalcOnLoad` in each file, which asks Excel to recalculate every
   formula on opening (the sources are quoted in `WorkbookWriter.cs`), so the formula results should be Excel's own. Row 10 of
   01-values shows whether it did.
3. Close **without saving**. Excel may ask whether to save changes, perhaps because it recalculated (not verified). Answer
   *Don't Save*.

Every file's last sheet, **Check**, says in one line what the file tests. It is also a sheet the edits did not touch, so it must look
as described: A1 `Check`, A2 the line, A3 `What to look at: docs/notes/excel-check.md in the repository.`

Every sheet the app wrote has a **hidden column XFD** (the last column), holding the picture settings in rows 1–3, such as
`ImageHeight:180`. That is expected (spec §5.2), and it is not part of the check.

Values were typed as in the app with the English (United States) language: for example, `3/15/2024` means 15 March 2024. Excel shows
dates and numbers in the PC's own format, so the expected values below are given without a format, as in "15 March 2024".

### 00-control.xlsx: open this one first

This file was made by the test builders alone (`TestWorkbook` in the tests): the app neither read nor wrote it. It is the source of
03-cards before the edits, with a picture added. **If Excel offers to repair this file, the test builders are at fault, not the
writer**, and the other files' repair prompts tell nothing about the writer until that is fixed. Note it in the results and stop.

Expected: sheets **Cards**, **Report**, **Check**. Cards: A1:B6 `Name`, `Qty`, Apple 1, Banana 2, Cherry 3, Damson 4, Elder 5; A20
`Total`, B20 **15**; D20:E20 merged; A24 on a yellow fill; a picture (number 0) at G2. Report: B2 **15**, B3 **1**, B4 **15**.

### 01-values.xlsx: typed values

Sheet **Values**. Column A says what was done to the cell beside it in column B. The type column says what kind of value the cell
must hold. To see it, type `=ISNUMBER(B2)`, `=ISLOGICAL(B2)` or `=ISTEXT(B2)` into a spare cell and fill it down. A date or a time
is a number to Excel, so `ISNUMBER` gives TRUE for it. Close without saving afterwards.

| Row | A | Expected in B | Type |
|---|---|---|---|
| 2 | Decimal typed | 3.5 | number |
| 3 | Whole number typed | 42 | number |
| 4 | Date typed | 15 March 2024 | date (number) |
| 5 | Date and time typed | 15 March 2024 14:30 | date (number) |
| 6 | Time typed | 14:30 (2:30 PM) | time (number) |
| 7 | TRUE typed | `TRUE` | boolean |
| 8 | false typed | `FALSE` | boolean |
| 9 | Formula typed | **7**; the formula bar shows `=B2*2` | formula |
| 10 | Formula kept | **43**; the formula bar shows `=B3+1`. The file still holds the old result, 11, so 43 shows that Excel recalculated | formula |
| 11 | Digits typed in a Text cell | `007`. The cell's number format is Text. A "number stored as text" marker is fine | text |
| 12 | Leading zeros typed | `007` | text |
| 13 | Words with a number typed | `12 apples` | text |
| 14 | Date kept | 15 March 2023, left as it was read (serial 45000) | date (number) |
| 15 | Number kept | 1234.5 | number |
| 16 | Boolean kept | `TRUE` | boolean |
| 17 | Text kept | `plain words` | text |
| 18 | Cleared | empty (it held `old text`); `=ISBLANK(B18)` gives TRUE | empty |

### 02-styles.xlsx: colours, fonts and alignment

Sheet **Styles**. All the texts in column B are in the default font, Calibri 11, unless the row says otherwise.

| Row | A | Expected in B |
|---|---|---|
| 2 | Fill set to yellow | `yellow fill` on a yellow fill (FFFF00) |
| 3 | Green fill removed | `no fill now`, no fill (white, with the grid lines showing) |
| 4 | Text colour set to blue | `blue text` in blue (0046C8) |
| 5 | Red text set to Automatic | `automatic text` in the automatic colour, not red. The file gives it no font colour. In Format Cells, Font, the Color box should say Automatic, not black (how Excel shows it is not verified) |
| 6 | Bold italic, text edited | `still bold italic`, bold and italic |
| 7 | Orange bold, untouched | `untouched`, bold, on an orange fill (FFC000) |
| 8 | Multi-line text typed | three lines, `line one`, `line two`, `line three`, wrapped in the cell and aligned to the top. The row may need to be taller to show all three; Excel may not grow it by itself |
| 9 | Centred, text edited | `still centred`, centred horizontally |

### 03-cards.xlsx: cards added, removed and reordered

The source table on sheet **Cards** had five cards in rows 2–6: Apple 1, Banana 2, Cherry 3, Damson 4, Elder 5. Below it, from row
20, there was content. The edits: Banana removed, Elder moved to the top, Fig 6 added at the end, Grape 7 inserted above Cherry. One
card fewer and two more, so everything below the table moved down one row.

Sheet **Cards**:
- A1:B1 `Name`, `Qty`. Rows 2–7: Elder 5, Apple 1, Grape 7, Cherry 3, Damson 4, Fig 6. Rows 8–20 are empty.
- A21 `Total`, B21 **20**, formula `=SUM(B2:B6)`. The range grew by the row inserted inside it (Grape) and shrank by the row removed
  (Banana). Fig was added below the range, so it is not counted, as when you insert a row under a range in Excel.
- A22 `Count`, B22 **5**, formula `=COUNTA(A2:A6)`.
- A23 `Cherry's qty`, B23 **1**, formula `=B3`. References to the table stay with the row, not the card, as after a sort (spec §9.20).
  Row 3 now holds Apple.
- A24 `Banana's qty`, B24 **#REF!**, formula `=#REF!`. Its row was deleted.
- D21:E21 merged into one cell, holding `merged note below the table`.
- A25 `yellow cell below the table` on a yellow fill.
- Page Layout, Print Area (or File, Print, preview): `A1:E25`.

Sheet **Report**: B2 **20** (`=Cards!B21`), B3 **5** (`=Cards!B2`, which is Elder's row now), B4 **20** (`=CardTotal`). In Formulas,
Name Manager, the name `CardTotal` refers to `=Cards!$B$21`.

### 04-table-name.xlsx: a table name inserted

The table on sheet **Named** had no name: header in row 1, three cards, a total in row 15. The table was given the name `Stationery`,
so a row was inserted above the header.

Sheet **Named**: A1 `Stationery`. Row 2 is the header, `Name`, `Price`. Rows 3–5: Pen 2, Ink 5, Pad 3. A16 `Total`, B16 **10**
(`=SUM(B3:B5)`). Sheet **Report**: B2 **10** (`=Named!B16`).

### 05-shared-formulas.xlsx: a shared formula split

Sheet **Shared** had `Double` = `=B2*2` filled down C2:C6 as one shared formula: the formula is stored once, in C2, for the whole range. The
edits: the card `Three` was removed, and `Two`'s Double was typed as `=B3*10`.

Sheet **Shared**: rows 2–5 are One 1 **2**, Two 2 **20**, Four 4 **8**, Five 5 **10**. Click each C cell: the formula bar shows `=B2*2`,
`=B3*10`, `=B4*2` and `=B5*2`. No cell shows an error, a 0 or a stale value.

### 06-sheets.xlsx: sheets renamed, duplicated, removed and reordered

The source had the sheets Plants, Summary, Old and Check. The edits: Plants renamed to `Garden plants` (a name with a space, so formulas
need quotes), then duplicated; Old removed; Summary moved to the front.

- Tabs, in this order: **Summary**, **Garden plants**, **Garden plants (Copy)**, **Check**. There is no `Old`.
- **Summary**: B2 **30** (`='Garden plants'!B2`), B3 **45** (`='Garden plants'!B3`), B4 **30** (`=FernHeight`), B5 **75**
  (`=SUM('Garden plants'!B2:B3)`). Name Manager: `FernHeight` refers to `='Garden plants'!$B$2`.
- **Garden plants**: A1:B3 `Name`, `Height` / Fern 30 / Rose 45. G1 `Note beside the table`, G2 `kept in the copy too`. Column G is
  wide: width 30, about 215 px. Page Layout: Orientation **Landscape**; margins left and right 0.5", top and bottom 0.75".
- **Garden plants (Copy)**: the same, except that B2 is **99**. It has the same column G, landscape orientation and margins.

### 07-pictures.xlsx: pictures in their cells

Every picture written by the app shows **a white box with a round black ring and a black number**. The ring touches the box's two longer
sides. **No red may show anywhere**: the red bands in the picture files lie outside the part the app keeps, so red at an edge means a
wrong crop. **An oval ring** means a stretched picture. The pictures are small files (40–120 px) shown larger, so they look blocky. That
is expected.

Sheet **Pictures**. The pictures are 100 × 150 px for a tall picture and 200 × 150 px for a wide one. Up to 4 units fit per row,
where a tall picture is 1 unit and a wide one 2.
- Column B is about 200 px wide (width 27.86). Column C is about 400 px wide (width 56.43). Rows 2–6 and 9 are 112.5 pt (150 px)
  tall, row 7 is 225 pt (300 px), and row 8 is 342 pt (456 px).
- B2 `Tall, fits`: one tall picture (1), 100 × 150.
- B3 `Tall, cropped top and bottom`: one tall picture (2), 100 × 150. The file is taller, with red bands above and below that must not
  show.
- B4 `Wide, fits`: one wide picture (3), 200 × 150.
- B5 `Wide, cropped at the sides`: one wide picture (4), 200 × 150, with no red at the left or right.
- B6 `Square, cropped top and bottom`: one wide picture (5), 200 × 150. A square counts as wide.
- C7 `Several in one cell`: five pictures in two rows, filling the cell's width. Top row: 1 (tall), 2 (wide), 3 (tall). Bottom row:
  4 (tall), 5 (wide). There is a gap at the right end of the bottom row.
- C8 `Thirteen, the card grows`: pictures 1–13, all tall, four per row, in order: 1–4, 5–8, 9–12 in row 8, and 13 alone at the left of
  row 9. The card now covers rows 8 and 9: A8:A9, B8:B9 and C8:C9 are each merged. Row 9 is the row the save inserted.
- Row 10 `Picture from the file removed`: B10 is empty. The source had a picture (number 99) there, which was removed.
- H2: a picture that was in the source outside the table, number **0**. It is kept as it was, filling cell H2 (anchored from H2's top
  left corner to I3's). Its look is not part of the check, only that it is still there.
- A22 `Below the table`, B22 **8** (`=COUNTA(A2:A10)`). It was in row 21. It moved down one row with the inserted row, and the range
  grew with it.

Sheet **Big pictures**. Its pictures are 200 × 300 px, one per row. That is taller than three to a worksheet row would allow (spec
§9.26), so each worksheet row holds one.
- B2 and B3: pictures 1 and 2, one in each row, each 200 × 300. Rows 2 and 3 are 229.5 pt tall. A2:A3 and B2:B3 are merged, with
  `Two tall pictures` in A2.
- A4 `After`. It was in row 3.

### 08-new-workbook.xlsx: a workbook made from scratch

Made in the app's model with no source file, as File, New does.
- Tabs: **Sheet1**, **Check**.
- **Sheet1**: A1 `Made from scratch`. Row 2 is the header, `Name`, `Count`, `Picture`. Row 3: `First`, the number **12**, and in C3 one
  picture (1), 136 × 180 px (the default size), with a round ring and no red. Row 4: `Second`, B4 **13** (`=B3+1`). Row 3 is 135 pt
  tall, and column C is about 136 px wide.
- **Check**: A1 `Check`, A2 `08 new workbook: made in the app's model with no source file.`

### 09-jpeg-orientation.xlsx: does Excel apply EXIF orientation? (spike/jpeg-orientation)

Sheet **Orientation** holds the same small made-up JPEG twice, each shown 200 × 100 px. The file is 96 × 48 px: its **top half is
red, its bottom half blue, with a black square at the top-left corner**.
- B2 `No EXIF`: the plain file. Expected: a red top half, a blue bottom half, and a black square at the top left.
- B3 `EXIF orientation 6`: the same pixels, with an EXIF tag that says "turn this a quarter turn clockwise to show it". If Excel
  applies the tag, the picture looks turned: **red on the right, blue on the left, the black square at the top right**, probably
  squeezed into the wide box. If Excel ignores the tag, B3 looks exactly like B2.

**The question for the spike:** does B3 look like B2 (Excel ignores EXIF orientation) or turned (Excel applies it)? Write the answer
in the results table below. If Excel applies it, `feat/picture-facts` should swap a JPEG's width and height for orientations 5–8, as
the roadmap says.

## Results

Filled in by the local check: the date, the Excel version (File, Account, About Excel), and per file whether it opened without a repair
prompt, whether everything listed matches, and what did not.

Excel version: … · Date: … · Windows display language and region format: …

| File | Opens without repair | Matches the list | Notes (what differed) |
|---|---|---|---|
| 00-control.xlsx | | | |
| 01-values.xlsx | | | |
| 02-styles.xlsx | | | |
| 03-cards.xlsx | | | |
| 04-table-name.xlsx | | | |
| 05-shared-formulas.xlsx | | | |
| 06-sheets.xlsx | | | |
| 07-pictures.xlsx | | | |
| 08-new-workbook.xlsx | | | |
| 09-jpeg-orientation.xlsx | | | B3: looks like B2 / turned |
