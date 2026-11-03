# App document check (PR #28)

The hands-on check of `feat/app-document` (PR #28): New, Open, Save and Save as with the system pickers, the unsaved-changes
question, error dialogs, drag and drop, the File menu's accelerators and access keys, and the exe's file description. It follows
the "Local check card" in the PR's description, steps 1–15. Step 16 (Narrator) is deferred to the milestone accessibility pass.

## Setup

- **Date:** 2026-11-03.
- **Windows:** Windows 11 Pro, version 26H2, build 26300.9711, x64. Display language English (United States).
- **Build:** PR #28, head `3ffb808`, CI run 41654845625 (`pull_request`, success), artifact **SpreadsheetEditor-win-x64**, unzipped
  to a scratch folder and run as is. The exe's file description reads `Spreadsheet Editor`, product version
  `0.1.0+7413dd444d30f87a3de31eac64777387e6a3d5c0` (the merge commit CI built).
- **Files:** made up only.
  - `.xlsx`: the Excel check samples, made as `excel-check.md` says, in `%TEMP%\excel-check`.
    `00-control.xlsx` (sheets Cards, Report, Check) served as the workbook with several sheets.
  - `.csv` and damaged files in `%TEMP%\app-document-check`:
    - `plants.csv` and `plants-semi.csv`: UTF-8 without BOM, CRLF, no line break after the last row, as Notepad saves them.
    - `empty.csv`.
    - `damaged.xlsx`: the text `hello`.
    - `old-renamed.xlsx` and `old.xls`: an OLE compound-file signature followed by zeros, not a real `.xls`.
    - `notes.txt`, a folder `a-folder`.
    - `big.csv`: 300,000 rows, 22 MB. `huge.csv`: 3,000,000 rows, 229 MB.
- **Who did what:** Claude drove the app through screenshots, clicks and keys, and read files and window titles back from the
  shell. The owner did step 11 (drag and drop from Explorer) and looked at the taskbar and Task Manager in step 1.
- **Not available:** desktop Excel is not installed on this PC. For step 12's "open in Excel", another process held `copy.xlsx` open
  for reading and writing with only read sharing allowed, as Excel does for a workbook it has open.

## Results

| Step | Result | Notes |
|---|---|---|
| 1. Start | Passed | Title `Spreadsheet Editor`. File shows New, Open, Save and Save as (both disabled), a separator and Exit, with Ctrl+N, Ctrl+O, Ctrl+S and Ctrl+Shift+S. The owner saw "Spreadsheet Editor" on the taskbar and in Task Manager. |
| 2. Access keys | Passed | Alt shows `F` on File. F opens the menu with N, O and X; S and A show once a document is open (their items are disabled before). Esc closes the menu and leaves key tips on File. Russian: see step 15. |
| 3. Open .xlsx | Passed | The picker hides `.txt` and `.xls` files. Its type list reads `*.xlsx`, `*.csv` and `All Files (*.xlsx;*.csv)`, see remark 1. `01-values.xlsx` shows "Values — cards: 17" and "Check — cards: 2". The title is `01-values.xlsx - Spreadsheet Editor`, and Save and Save as are enabled. |
| 4. Save | Passed | Ctrl+S rewrote `01-values.xlsx` (new modified time) with no picker. It opens again with the same lines. |
| 5. Save as | Passed | The picker opens in `%TEMP%\excel-check` with "Excel workbook (*.xlsx)" first and "CSV file (*.csv)" second. The name box reads `01-values.xlsx`, see remark 2. Typing `copy` saved `copy.xlsx`, and the title became `copy.xlsx - Spreadsheet Editor`. |
| 6. Save as .csv with several sheets | **Failed** | `00-control.xlsx` saved as `multi` with "CSV file" wrote `multi.csv` with the first sheet only (`Name,Qty` and five rows), and Ctrl+S wrote it again with no picker. **But the title reads `multi.csv - Spreadsheet Editor`, without the `*`.** The document was never edited, so it stays unmodified, see "Failure" below. |
| 7. CSV round trip | Passed | `plants-semi.csv` after Ctrl+S: still `;`, still UTF-8 without BOM, CRLF, the same three lines. One difference: a line break was added after the last row, see remark 6. |
| 8. New | Passed | `Untitled - Spreadsheet Editor`, "Sheet1 — cards: 1". Ctrl+S opens the save picker with `Untitled.xlsx` in the last folder used. Cancel left everything as it was. Russian (Alt, Ф, С): `Без имени - Spreadsheet Editor`, "Лист1 — карточек: 1". |
| 9. Unsaved-changes question | Not run | It needs a modified document. This build has no editing, and step 6 does not make one (see "Failure"). With an unmodified document, Alt+F4 closed the app without a question, as it should. |
| 10. Untitled save from the question | Not run | As step 9. |
| 11. Drop | Passed | Done by the owner: `plants.csv` shows the Copy cursor and opens. `notes.txt`, `old.xls` and a folder show no-drop, and nothing happens. Of several files with `plants.csv` first, only it opens. |
| 12. Error dialogs | Passed | Each dialog's title and message are as the card gives them. The previous document stays open, and the app keeps running. See remarks 3–5. <br>• `damaged.xlsx`: "Couldn't open damaged.xlsx", "This file isn't an Excel workbook (.xlsx)." The Details text can be selected and copied: the clipboard held "The file is not an Excel workbook (.xlsx)." <br>• `empty.csv`: "The file is empty: it has no header line." <br>• `old-renamed.xlsx`: the older-workbook or password message. <br>• Read-only `copy.xlsx`, Ctrl+S: "Couldn't save copy.xlsx", the access message, and the modified time did not change. <br>• `copy.xlsx` locked, Ctrl+S: the same access message, and the file did not change. No temporary files were left in the folder. |
| 13. Pickers remember the folder | Passed | After opening from `%TEMP%\excel-check`, the next Open and Save as start there. After a restart, Open starts in Documents, which the card allows. |
| 14. Busy state | Passed | `big.csv` (300,000 rows) loads in under 0.3 s, too fast to see. `huge.csv` (3,000,000 rows) loads in a few seconds: the progress bar shows, New, Open, Save, Save as and Exit are disabled, and the window can be dragged while it loads. Then "huge.csv — cards: 3000000". |
| 15. Russian | Passed (steps 1, 2, 12); step 9 not run | `settings.json` = `{ "language": "ru" }`, removed afterwards. <br>• Menu: Файл with Создать, Открыть, Сохранить, Сохранить как and Выход, and the accelerators. <br>• Key tips: Ф, then С, О, Х, К and В. <br>• Error titles read `Не удалось открыть «damaged.xlsx»` and `Не удалось сохранить «copy.xlsx»`, with Подробности. The messages for `damaged.xlsx`, `empty.csv`, `old-renamed.xlsx` and read-only `copy.xlsx` are in Russian. <br>• No text was cut off in the menu or the dialogs. <br>• The Details text stays English, see remark 5. Step 9 could not be reached, as in English. |

## Failure

**Step 6: a CSV save that leaves sheets unwritten does not mark an unedited document modified.**

- **Expected (the card):** `*multi.csv - Spreadsheet Editor`. Spec §9.34 (decided 2026-11-03) says "the document stays modified when
  it has other sheets that were not written".
- **Seen:** `multi.csv - Spreadsheet Editor`. The title and the window title (as the shell reads it) both lack the `*`.
- **Why, read in the branch's code, not changed:**
  - `MainViewModel.SaveToAsync` does not call `MarkSaved` when `SaveResult.AllSheetsWritten` is false.
  - `UndoHistory.IsModified` is `undo.Count != savedDepth`.
  - A document that was opened and never edited is therefore unmodified before the save and stays so after it.
- **For the manager:** decide whether such a save must mark the document modified even with no edits, as the card expects, or only
  keep an edited document modified. If the latter, the card's step 6 and step 9's way of getting a modified document need changing.
  As the build stands, there is no way to get a modified document, so steps 9 and 10 (and step 9 in Russian) remain to be checked.

## Remarks (not failures)

1. **Open picker type names.** The open picker's types read `*.xlsx`, `*.csv` and `All Files (*.xlsx;*.csv)` (ru `Все файлы (…)`),
   with the combined one chosen. They have no names such as "Excel workbook", unlike the save picker. "All Files" is misleading,
   since it holds only the two types.
2. **Save as name with its extension.** The save picker's name box shows `01-values.xlsx`, not `01-values`. Explorer shows file
   name extensions on this PC, and the picker then shows the extension (not verified as the cause). Typing a name without an
   extension (`copy`, `multi`) saved with the chosen type's extension, so the PR's open point about that is answered: it does.
3. **Error dialog focus.** The dialog opens with focus on **Details**, so Enter expands Details rather than choosing OK. Esc and
   clicking OK close it.
4. **A locked file gets the read-only message.** Saving over a file another program holds open reports "Windows doesn't allow access
   to this file. Check that the file isn't read-only and that you have permission to use it." The Details read "The file cannot be
   written: Access to the path is denied." The message does not suggest that another program has the file open. Checked with a
   simulated lock, not with Excel.
5. **Details text is not localised.** In Russian, the Details text is the core's English exception message, for example "The file is
   not an Excel workbook (.xlsx)."
6. **CSV final line break.** Saving `plants-semi.csv` added a CRLF after the last row, which the file did not have. The rows are
   otherwise byte for byte the same.
7. **Ctrl+S behind a dialog.** While an error dialog is open, Ctrl+S does nothing, as it should: the file's modified time did not
   change.
8. **Folder after a failed open.** After a failed open, the next Open starts in the folder of the last successful open, not that of
   the file that failed.
