# Spreadsheet Editor: specification

Spreadsheet Editor is a Windows desktop app for editing a table in a spreadsheet as a stack of **cards**: one card per table row, one **entry** per table column. Each entry has text, an optional fill colour, a text colour and any number of **pictures**. When the app saves to `.xlsx` it lays the pictures out in a grid inside the cells: it sizes columns and rows, inserts and merges rows where needed, and crops pictures to a fixed aspect ratio. The file can then be printed or shared like any other workbook.

This document covers what the app does, the file formats it reads and writes, and every message it shows. Sections 1 to 8 describe the behaviour to build. Section 9 lists deliberate changes and open questions: build each "Intended" and "Decided" statement in place of what sections 1 to 8 say, and keep sections 1 to 8 for each "Q" until the owner answers it. Section 10 lists where sections 1 to 8 depart from Windows 11 conventions; the owner decided to follow Windows there (ADR 0001), so the Windows alternative in section 10 replaces what sections 1 to 8 say. When an owner answers a question, the answer replaces the question here, with its date.

Terms:

| Term | Meaning |
|---|---|
| Workbook | The open document: one or more sheets. |
| Sheet | One worksheet, shown as one tab. Each sheet holds exactly one table. |
| Table | The block of cells on the worksheet that the app edits (section 4.1). |
| Entry | A table column. It has a name (the header), which may be empty. |
| Card | A table row, or a group of vertically merged worksheet rows (section 4.3). |
| Value | The content of one entry in one card. |
| Picture | An embedded image that belongs to one value. |
| Picture unit | A horizontal slot one picture-width wide. A picture is **wide** if its pixel width is at least its pixel height; wide pictures take 2 units. Other pictures are **tall** and take 1 unit. |

Product name: **Spreadsheet Editor**. The app ships as a self-contained folder (later an installer) with the application icon in `src/SpreadsheetEditor.App/Assets`.

---

## 1. Concept and layout

### 1.1 Window

- Initial size: 1280 × 800. No minimum size is enforced.
- Title: `Spreadsheet Editor`, or `Spreadsheet Editor *` while there are unsaved changes. The title is never translated.
- The window has the application icon.
- From top to bottom, the window contains the top bar and then the sheet tab area. The tab area fills the remaining space.

### 1.2 Top bar

Left to right:

1. **Upload Spreadsheet** button: opens a file (section 2.1).
2. **Save Spreadsheet** button: saves (section 2.2). It is disabled at startup. It becomes enabled once a file has loaded (even one with no visible sheets) or a sheet has been added. It becomes disabled again when the last sheet is removed.
3. A flexible space.
4. **Language selector** (drop-down), right-aligned. It lists `English` and `Russian`, shown in the current UI language (section 7.2).

### 1.3 Sheet tabs

- There is one tab per sheet, and the tab text is the sheet name.
- Each tab has its own close button (an "×" drawn in the current text colour). Pointing at it gives it the danger background colour.
- Tabs can be dragged to reorder them.
- A round **"+"** button (24 × 24, tooltip `Add Sheet`) floats 4 px to the right of the last tab and is centred vertically in the tab strip. When there are no tabs it sits 4 px from the left edge of the strip. It moves whenever tabs are added, removed, renamed or moved, or the strip is resized.
- Tab styling: padding 5 px top and bottom, 10 px left, 25 px right. Top corners have a 5 px radius, with a 2 px gap between tabs. The selected tab uses the card background and bold text. A pointed-at tab uses the hover background. The tab content pane has a 1 px border.
- At startup there are no tabs, and the content pane is empty.

### 1.4 Sheet page (one per tab)

The settings bar is a single horizontal row at the top. Below it, the card list scrolls vertically, with cards stacked from the top.

Settings bar, left to right:

| Control | Details |
|---|---|
| Label `Table Name:` | |
| Table name field | Single line, minimum width 150. Placeholder: `No table name`. Starts with the sheet's table name, or empty. |
| `Add Card` button | Adds an empty card at the end. |
| `Add Entry` button | Adds an entry at the end. |
| `Remove Entry` button | Opens the Remove Entries dialog. |
| `Reorder Entries` button | Opens the Reorder Entries dialog. |
| Picture width box | Integer, range 50–1000. Shown as `Image Width: <n>px`. Default 136. |
| Aspect-lock toggle | A toggle button 35 px wide, showing a link symbol (🔗). Tooltip `Lock Aspect Ratio`. Off by default. |
| Picture height box | Integer, range 50–500. Shown as `Image Height: <n>px`. Default 180. |
| Pictures-per-row box | Integer, range 1 to max(1, floor(1762 / picture width)). Shown as `Images per row: <n>`. Default 8. |

The card list area uses the scroll-area background colour (section 7.1).

### 1.5 Card

A card is a framed panel: 5 px outer margin, 5 px padding, the card background, a 1 px border and a 5 px corner radius. Inside it are a two-column form with one row per entry, followed by a bottom row.

**Each entry row:**

- **Label column.** It holds an editable entry name field, 150 px wide, with placeholder `Unnamed`. Unnamed entries show an empty field. The field is absent when the sheet's headers are hidden (section 3.3); the value then spans the row.
- **Value column**, made of:
  - A top line containing:
    - A **multi-line text field**, which takes all remaining width. Its height fits the content: at least one line (line spacing + 14 px), growing with wrapped lines up to 70 px. At 70 px a vertical scroll bar appears as needed; below 70 px there is never a scroll bar. The height is recalculated when the text changes and when the field is resized.
    - An **`Add Attachment`** button, aligned to the top.
    - A **fill swatch**: 35 × 24, tooltip `Background Color`.
    - A **text-colour swatch**: 35 × 24, tooltip `Text Color`.
  - Below the top line, the **picture strip**. It appears only when the value has at least one picture, and is described in section 6.2.

  Swatch appearance:
  - With a colour: filled with that colour, 1 px solid #888 border, 2 px margin.
  - Without a colour: transparent, with a 1 px dashed #888 border.

**Bottom row:**

- `↑` button (30 px wide).
- `↓` button (30 px wide).
- Flexible space.
- `Remove Card` button.

The text fields do **not** show the value's fill, text colour, font or alignment. Only the swatches show colours.

---

## 2. User actions

### 2.1 Opening a file

There are three ways to open a file:

- **Upload Spreadsheet** button.
- Drag one or more files onto the window. The drag is accepted if any dragged file ends in `.xlsx`, `.xls` or `.csv` (case-insensitive). On drop, only the **first** file is used. If the first file is not one of those types, nothing happens.
- **Ctrl+V** when the keyboard focus is not in a text field or picture strip. If the clipboard holds file paths and the first one ends in `.xlsx`, `.xls` or `.csv`, that file is opened. Otherwise nothing happens.

Flow:

1. If there are unsaved changes **and** at least one sheet is open, show a confirmation:
   - Title: `Unsaved Changes`
   - Text: `Loading a new spreadsheet will lose any unsaved changes. Continue?`
   - Buttons: Yes / No.

   Anything other than Yes cancels. For the button, this confirmation comes **before** the file dialog.
2. The button shows the open-file dialog:
   - Title: `Open Spreadsheet`
   - Filter: `Spreadsheets (*.xlsx *.xls *.csv)`
   - Single selection, no initial folder.

   Cancelling the dialog does nothing.
3. Remove all tabs, then load the file (sections 4 and 5.8).
4. On success:
   - One tab is created per loaded sheet, in workbook order, and the first tab is selected.
   - Save is enabled.
   - The document is marked unmodified.
5. On failure, show an error:
   - Title: `Error`
   - Text: `Could not load spreadsheet: {error}`, where `{error}` is the underlying error message.

### 2.2 Saving

Triggered by **Save Spreadsheet**. If no sheets are open, nothing happens.

1. Show the save-file dialog:
   - Title: `Save Spreadsheet`
   - Filters: `Excel (*.xlsx)` and `CSV (*.csv)`
   - No initial folder or file name.

   Every save asks for a path; there is no separate "save to the current file". Cancelling does nothing.
2. If the chosen path ends in `.csv` (case-sensitive), export CSV (section 5.9). Otherwise write `.xlsx` (section 5).
3. On success:
   - Mark the document unmodified.
   - Show a message: title `Success`, text `Spreadsheet saved successfully!`.
4. On failure:
   - Show an error: title `Error`, text `Could not save spreadsheet: {error}`.
   - The modified state is unchanged.

The values written are the current state of every sheet in tab order:

- name (tab text)
- entries and cards
- table name (the table name field's text)
- table name column and header row
- hidden-headers flag
- picture width, picture height and pictures per row (from the three boxes)
- first column

### 2.3 Sheets

| Action | How | Behaviour |
|---|---|---|
| Add sheet | "+" button | Shows a text-input dialog (title `Add Sheet`, label `Sheet Name:`, OK/Cancel). If the user cancels or leaves the name empty, nothing happens. If another tab already has exactly this name (case-sensitive), shows a warning (title `Warning`, text `Sheet '{name}' already exists.`) and stops. Otherwise it appends a sheet (see the defaults after this table), selects its tab, enables Save and marks the document modified. |
| Rename sheet | Double-click a tab, or tab context menu → `Rename` | Shows a text-input dialog (title `Rename Sheet`, label `New Sheet Name:`, pre-filled with the current name). If the user cancels, leaves it empty or leaves it unchanged, nothing happens. If another tab has the same name, shows a warning (title `Warning`, text `Sheet '{new_name}' already exists.`). Otherwise it renames the tab and marks the document modified. |
| Duplicate sheet | Tab context menu → `Duplicate` | Inserts a copy directly after the source tab and selects it. The name is `<name> (Copy)`. If that name is taken, it tries `<name> (Copy 1)`, `<name> (Copy 2)` and so on. The suffix is never translated. The copy takes the source's current entries, cards (with all values, styles and pictures), row spans, table name column, header row, hidden-headers flag, first column, picture width/height and pictures per row. It does **not** copy the aspect lock or the table name as edited: it uses the table name the sheet had when it was loaded (see section 9). Marks the document modified. |
| Remove sheet | Tab close button, or tab context menu → `Remove` | Shows a confirmation (title `Confirm Remove`, text `Are you sure you want to remove sheet '{name}'?`, Yes/No). On Yes it removes the tab and marks the document modified. If no tabs remain, it marks the document **unmodified** and disables Save. |
| Reorder sheets | Drag a tab | Moves the tab. The document is **not** marked modified (see section 9). |

A new sheet has these defaults:

- one unnamed entry and one card with empty text
- headers shown, no table name
- table name column 1, header row 0, first column 1
- picture width 136, picture height 180, 8 pictures per row

The tab context menu opens on right-click on a tab. It contains `Duplicate`, `Rename` and `Remove`. Right-clicking the empty part of the tab strip shows no menu.

### 2.4 Cards

| Action | How | Behaviour |
|---|---|---|
| Add card at end | `Add Card` | Appends an empty card. Every value is empty text. Picture entries (section 3.4) get an empty picture value. |
| New card above / below | Card context menu → `New card above` / `New card below` | Inserts an empty card before or after this card. |
| Duplicate card | Card context menu → `Duplicate` | Inserts directly below a deep copy of this card: text, colours, pictures, alignment, font and all other style data. |
| Remove card | `Remove Card`, or card context menu → `Remove` | Removes the card immediately, without confirmation. If it is the only card, shows a warning instead (title `Warning`, text `Cannot remove the last card.`). |
| Move up / down | `↑` / `↓` | Swaps the card with its neighbour. On the first card (`↑`) or last card (`↓`) nothing happens; the buttons stay enabled. |

The card context menu opens on right-click anywhere on the card except text fields, which have their own menu, and the picture strip. Its items, in order:

- `New card above`
- `New card below`
- `Duplicate`
- separator
- `Remove`

Every card action marks the document modified.

### 2.5 Entries

| Action | How | Behaviour |
|---|---|---|
| Add entry | `Add Entry` | Appends an unnamed entry to every card, with empty text. |
| Rename entry | Type in an entry name field in **any** card | Each change renames the entry immediately. The name field of that entry in every other card updates to match. |
| Remove entries | `Remove Entry` | Opens a dialog (see below). |
| Reorder entries | `Reorder Entries` | Opens a dialog (see below). |

**Remove Entries dialog:**

- Title: `Remove Entries - {name}`, where `{name}` is the sheet name. Size: 300 × 400.
- Label: `Check entries to remove:`.
- A list with one checkbox item per entry, in order, all unchecked. Item text is the entry name, or `Unnamed` for an unnamed entry.
- Buttons `OK` and `Cancel`, side by side in that order.
- On OK with at least one item checked, the checked entries are removed from every card, without further confirmation. Nothing stops the user from removing every entry (see section 9).

**Reorder Entries dialog:**

- Title: `Reorder Entries - {name}`. Size: 300 × 400.
- Label: `Drag and drop to reorder entries:`.
- A list of entry names (`Unnamed` for unnamed entries) that can be reordered by drag and drop.
- Buttons `OK` and `Cancel`.
- On OK, the new order is applied to every card.

All entry actions mark the document modified.

### 2.6 Values

- **Text:** edit in the multi-line field.
- **Fill colour:** click the fill swatch. A small menu opens below the swatch with `Pick Color...` and `No Fill`:
  - `No Fill` clears the fill.
  - `Pick Color...` opens the colour dialog. It starts on the current fill, or white if there is none. Choosing a colour sets the fill to `#RRGGBB` (upper case).
  - Dismissing the menu or the dialog changes nothing.
- **Text colour:** click the text-colour swatch. The colour dialog opens directly (no menu), starting on the current text colour, or black if there is none. There is no way to clear the text colour.
- **Pictures:** see section 6.

Every change marks the document modified.

### 2.7 Table name

Editing the table name field marks the document modified.

### 2.8 Picture size settings

- **Picture width changed:**
  - All cards are told the new width.
  - If the aspect lock is on, height = round(width / ratio), clamped to 50–500. The height box is updated without triggering its own change handling.
  - The pictures-per-row maximum becomes max(1, floor(1762 / width)). A current value above the new maximum is reduced to it.
  - Marks modified.
- **Picture height changed:**
  - If the aspect lock is on, width = round(height × ratio), clamped to 50–1000, and the pictures-per-row maximum is recomputed the same way.
  - Marks modified.
- **Aspect lock:** turning it on captures ratio = width / height at that moment. The ratio is not recaptured after clamping. Turning it off just stops the coupling. The lock is not saved anywhere.
- **Pictures per row changed:** marks modified.

Changing a value to the value it already has does nothing.

These three settings affect only the saved `.xlsx`. In-app previews are always 100 × 100 (section 6.2).

### 2.9 Keyboard

| Where | Key | Effect |
|---|---|---|
| Window (focus not in a field or picture strip) | Ctrl+V | Open a spreadsheet file from the clipboard (section 2.1). |
| Entry text field | Enter | Move focus to the next value text field (see "Focus movement" below). No newline is inserted. |
| Entry text field | Shift+Enter | Insert a newline. |
| Entry text field | Ctrl+V | If the clipboard holds image data, or file paths of which at least one has a picture extension, add the pictures to this value (section 6.1) instead of pasting text. Otherwise paste text normally. |
| Entry text field | Tab | Inserts a tab character (it does not move focus). |
| Entry name field, table name field | Enter | Move focus to the next value text field. |
| Entry name field, table name field | Ctrl+V | Normal text paste. |
| Picture strip (when focused) | Ctrl+V | Add pictures from the clipboard (section 6.1). |
| Any text field | Focus gained by Tab, Shift+Tab, window activation or programmatically | Caret goes to the end of the text. |

Focus movement: "next value text field" means the next text field in tab order that is visible, enabled and focusable, and is not an entry name field or the table name field.

There are no application-wide shortcuts for open, save, new sheet and so on, and no application-level undo or redo. Text fields keep their built-in text undo.

### 2.10 Drag and drop summary

- **Spreadsheet file onto the window:** opens it (section 2.1).
- **Image files or image data onto a picture strip or an entry text field:** adds the pictures to that value (section 6.1). Other drops onto a text field behave as normal text drops.
- **Tabs:** drag to reorder.
- **Reorder Entries dialog:** drag items to reorder.
- **Cards:** cannot be dragged. Use `↑` and `↓`.

### 2.11 Context menus (complete list)

| Target | Items | Enabled when |
|---|---|---|
| Sheet tab | `Duplicate`, `Rename`, `Remove` | always |
| Card | `New card above`, `New card below`, `Duplicate`, —, `Remove` | always |
| Fill swatch (left-click) | `Pick Color...`, `No Fill` | always |
| Picture thumbnail | `Copy selected photo` | always |
| Picture strip background | `Copy all photos`, `Paste` | Copy all: the value has at least one picture. Paste: the clipboard has image data or file paths. |
| Text fields | The standard text-editing menu | standard |

### 2.12 Confirmations, warnings, errors and messages (complete list)

| Kind | Title | Text | Buttons |
|---|---|---|---|
| Confirm | `Unsaved Changes` | `Loading a new spreadsheet will lose any unsaved changes. Continue?` | Yes / No |
| Confirm | `Unsaved Changes` | `Any unsaved changes will be lost. Are you sure you want to exit?` | Yes / No |
| Confirm | `Confirm Remove` | `Are you sure you want to remove sheet '{name}'?` | Yes / No |
| Warning | `Warning` | `Sheet '{name}' already exists.` (add) / `Sheet '{new_name}' already exists.` (rename) | OK |
| Warning | `Warning` | `Cannot remove the last card.` | OK |
| Error | `Error` | `Could not load spreadsheet: {error}` | OK |
| Error | `Error` | `Could not save spreadsheet: {error}` | OK |
| Information | `Success` | `Spreadsheet saved successfully!` | OK |

Closing the window shows the exit confirmation only if there are unsaved changes **and** at least one sheet is open. Anything other than Yes keeps the window open.

---

## 3. Data model

### 3.1 Workbook

The workbook holds the following:

- An ordered list of sheets. Tab order is the order used when saving.
- A **source workbook**: the `.xlsx` that was opened, or nothing if a CSV was opened or nothing has been opened. Saving to `.xlsx` writes into the source workbook so that everything outside the tables is kept (section 5).
- A **modified** flag (section 8).

### 3.2 Sheet

| Field | Type | Default (new sheet) | Meaning |
|---|---|---|---|
| name | text | user-entered | Worksheet name and tab text. |
| entries | ordered list | one unnamed entry | Table columns. |
| cards | ordered list | one empty card | Table rows. |
| tableName | text or none | none | Optional title written above the header row. |
| tableNameColumn | 1-based column | 1 | Column of the table name cell. |
| headerRow | 0-based offset H | 0 | Header cells are on worksheet row H+1 (1-based). The table name is on row H, or row 1 if H = 0. |
| headersHidden | bool | false | See section 3.3. |
| firstColumn | 1-based column | 1 | Worksheet column of the first entry. |
| pictureWidth | int px | 136 | Width of one picture unit in the saved file. |
| pictureHeight | int px | 180 | Height of one picture row in the saved file. |
| picturesPerRow | int | 8 | Maximum picture units per picture row in one cell. |
| aspectLocked | bool | false | UI state only; never saved. |
| pictureEntries | set of entries | empty | Entries that held a picture in any card when loaded (section 3.4). |

Data rows start on worksheet row H+2, or H+1 when the headers are hidden.

### 3.3 Entry

- An entry has a **name**. A name that is empty, or begins with `Unnamed:`, counts as **unnamed**. An unnamed entry shows an empty name field with placeholder `Unnamed`, is listed as `Unnamed` in the dialogs, and is written as an empty header cell.
- Internal placeholder names:
  - On load, an unnamed entry is named `Unnamed: {i}`, where i is its 0-based position in the table.
  - A new entry is named `Unnamed: {n}`, where n is the current number of entries, with `_` appended until it is unique.
- **Hidden headers:**
  - When the header row of a loaded table contains any numeric value (booleans count as numeric), the sheet has hidden headers.
  - The entries are identified only by position (named 0, 1, 2, …).
  - No entry name fields are shown, and the header row is read and written as data.
  - Entries added later get `Unnamed: {n}` names but still show no name field.
- Two entries may end up with the same name: the user can type one. Entries must therefore be identified internally by a stable identity, not by name (section 9).

### 3.4 Card and values

A card has:

- **rowSpan:** the number of worksheet rows the card occupied in the source (section 4.3). 1 for new cards and for anything not loaded from `.xlsx`.
- One **value** per entry.

A value is one of the following.

1. **Plain value:** text, number, date, boolean or empty. These are produced by CSV loading and by new cards.

2. **Styled value:**
   - `text` (string)
   - `fill`: `#RRGGBB`, or none
   - `textColor`: `#RRGGBB`. It defaults to `#000000` whenever it is missing.
   - `alignment`: `{ horizontal, vertical, wrapText }`, each optional, using the spreadsheet's own values, for example `left`, `center`, `right`, `top`, `bottom`
   - `font`: `{ name, size, bold, italic, underline, strike }`, each optional; underline is the spreadsheet's underline kind, such as `single` or `double`

   Any extra style data read from the file is carried along unchanged.

3. **Picture value:** a styled value plus:
   - `pictures`: an ordered list of the original image file bytes, never re-encoded unless they came from clipboard or drag image data (section 6.1)
   - `as_text`: a boolean that is always `false`. Nothing in the app reads it (see section 9).

   Pictures created for entries that had none in a given card have `fill` none and `textColor` `#000000`, and no alignment or font.

Every entry in the UI can hold pictures, not only picture entries. `pictureEntries` only decides whether empty values in new cards start as empty picture values.

**What a card returns** when the sheet is read back (for saving, duplicating, or entry operations):

- For each entry it builds `{pictures, text, fill, textColor}` from the widgets and adds every other key of the value as originally loaded, such as alignment, font and `as_text`.
- It returns plain text only when there are no pictures, no fill, no text colour and no other keys.
- The text-colour swatch always holds a colour (default `#000000`), so in practice **every value that passes through a card becomes a styled value** with an explicit text colour (see section 9).

---

## 4. Reading files

### 4.1 `.xlsx`: finding the table on each sheet

Repeat for every worksheet, in workbook order.

1. **Occupied cells.**
   - Collect every cell whose value is not empty.
   - Add the top-left anchor cell of every picture that has a cell anchor. Pictures with an absolute position are ignored.
   - If the sheet has no occupied cells it is **skipped**: no tab is created, and on save it is deleted from the workbook (see section 9).
2. **Row block.**
   - Sort the distinct occupied rows. The block starts at the smallest.
   - Walk upward through them. The block ends at the last row before the first jump of **more than 10 rows** between consecutive occupied rows. (Up to 9 empty rows in between is still the same block.)
   - The block is first_row … last_row.
3. **Column block.**
   - Take the distinct columns of occupied cells within the row block, sorted.
   - Start at the smallest. End at the last column before the first jump of **more than 3 columns** (up to 2 empty columns in between is still the same block).
   - The block is first_col … last_col, and firstColumn = first_col.
4. **Table name.**
   - Look at row first_row within first_col … last_col.
   - The row has a table name if all three of these hold:
     - it has exactly one non-blank cell (text that is only spaces counts as blank);
     - first_row < last_row;
     - row first_row+1 has **more than one** non-blank cell in the column block.
   - If so:
     - tableName = that cell's value as text
     - tableNameColumn = that cell's column
     - headerRow H = first_row, so the header is on first_row+1
   - Otherwise:
     - tableName = none
     - tableNameColumn = first_col
     - H = first_row − 1, so the header is on first_row
5. **Header.**
   - Read row H+1 across the column block.
   - If any non-blank header cell is a number (or boolean), headersHidden = true and the entries are named by position 0…n−1.
   - Otherwise each entry is named by its header cell as text, or `Unnamed: {i}` if blank.
6. **Pictures.**
   - Every anchored picture whose anchor column lies inside the column block is assigned to entry (anchor column − first_col) at its anchor row.
   - The picture's original bytes are kept.
   - Pictures anchored outside the column block stay in the workbook but are not shown.
7. **Cards.** See section 4.3, starting at worksheet row H+2 (H+1 when the headers are hidden) and continuing through last_row.
8. **Layout metadata.** See section 4.5.

### 4.2 Values

For each card, each entry's value comes from the cell in the card's **first** worksheet row.

1. **Dates.** A date or date-and-time cell becomes text in the **short date format of the current UI language**, for example `1/15/24` for English and `15.01.2024` for Russian. The time part is dropped, and the conversion happens once, at load. Time-only values become their plain text form, such as `12:30:00`.
2. **Colours and styles.**
   - **fill:** the cell's pattern foreground colour, resolved as in section 4.4. Fully transparent (`00000000`) or absent means none.
   - **textColor:** the font colour, resolved as in section 4.4. If unresolvable or absent, it is `#000000`.
   - **alignment:** horizontal, vertical and wrapText, as stored on the cell.
   - **font:** name, size, bold, italic, underline and strike, as stored on the cell.
3. **Value kind.**
   - If any picture is assigned to this entry in any of the card's worksheet rows, the value is a **picture value**:
     - pictures are in worksheet-row order, then in the order they appear in the file
     - text is the cell value as text, or empty if the cell is empty
     - `as_text` is false
     - plus the colours and styles
   - Otherwise, if there is a fill, a text colour, a horizontal or vertical alignment, or any font attribute, the value is a **styled value** with the cell value as text. Because textColor always has a value, every `.xlsx` cell takes this path.
   - Otherwise the value is a plain value.
4. **Picture entries.** Any entry that received a picture in at least one card becomes a picture entry. In every other card, that entry's value is converted to an empty picture value: text = the old value as text, fill none, textColor `#000000`.

Numbers are shown as their plain text form (`3`, `3.5`), and booleans as `True` / `False` (replaced by section 9.3: `TRUE` / `FALSE`).

A cell that holds a formula is shown as its formula text, beginning with `=`.

### 4.3 Grouping vertically merged rows into cards

```
row = dataStartRow
while row <= last_row:
    end = row
    repeat until nothing changes:
        for each merged range on the worksheet (any columns):
            if range overlaps rows row…end and range.lastRow > end:
                end = range.lastRow
    add card covering rows row…end   (rowSpan = end − row + 1)
    row = end + 1
```

- Merged ranges in **any** column extend a card, including columns outside the table.
- Values come only from the first row of the group (section 4.2), but pictures are collected from every row of the group.
- Empty worksheet rows inside the table become empty cards.

### 4.4 Colours

**Explicit colour (ARGB):**

- 8 hex digits: `00000000` means none. Otherwise the result is `#` followed by the last 6 digits; the alpha digits are ignored.
- 6 hex digits: `#` followed by the digits.

**Theme colour (index + optional tint):**

1. **Palette.** Read the workbook theme's colour scheme in this order: `lt1, dk1, lt2, dk2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink`.
   - Each colour comes from its sRGB value.
   - If it is a system colour, use its last-known colour, defaulting to `FFFFFF`.
   - A missing slot is `FFFFFF`.
   - If the theme cannot be read, use the default palette:

     | Index | Slot | Default |
     |---|---|---|
     | 0 | lt1 | FFFFFF |
     | 1 | dk1 | 000000 |
     | 2 | lt2 | E7E6E6 |
     | 3 | dk2 | 44546A |
     | 4 | accent1 | 5B9BD5 |
     | 5 | accent2 | ED7D31 |
     | 6 | accent3 | A5A5A5 |
     | 7 | accent4 | FFC000 |
     | 8 | accent5 | 4472C4 |
     | 9 | accent6 | 70AD47 |
     | 10 | hlink | 0563C1 |
     | 11 | folHlink | 954F72 |

2. **Index out of range** (outside 0–11): no colour.
3. **No tint, or tint 0:** the palette colour as is.
4. **Tint:** convert the colour to HSL with this algorithm (all channels in 0–1):
   ```
   r,g,b = R/255, G/255, B/255
   max = max(r,g,b); min = min(r,g,b); l = (max+min)/2
   if max == min: h = s = 0
   else:
     d = max − min
     s = l > 0.5 ? d / (2 − max − min) : d / (max + min)
     h = max==r ? (g − b)/d + (g < b ? 6 : 0)
       : max==g ? (b − r)/d + 2
       :          (r − g)/d + 4
     h = h / 6
   ```
   Apply the tint to the lightness:
   - tint < 0: `l = l × (1 + tint)`
   - tint > 0: `l = l × (1 − tint) + tint`

   Convert back to RGB:
   ```
   if s == 0: r = g = b = l
   else:
     q = l < 0.5 ? l × (1 + s) : l + s − l × s
     p = 2l − q
     r = hue(p,q,h + 1/3); g = hue(p,q,h); b = hue(p,q,h − 1/3)
   hue(p,q,t): if t<0: t+=1; if t>1: t−=1
               if t<1/6: return p + (q−p)·6t
               if t<1/2: return q
               if t<2/3: return p + (q−p)·(2/3 − t)·6
               return p
   channel = floor(x × 255 + 0.5)
   ```
   The result is `#RRGGBB` in upper case.

5. **Other colour kinds** (indexed or legacy palette, automatic): no colour. A text colour then falls back to `#000000`.

### 4.5 Layout metadata cells

Each sheet's picture settings are stored in hidden cells in the **last worksheet column (column 16384, `XFD`)**:

| Row | Text format | Default if missing or invalid |
|---|---|---|
| 1 | `ImageHeight:{int}` | 180 |
| 2 | `ExcelImagesPerRow:{int}` | 8 |
| 3 | `ImageWidth:{int}` | 136 |

- A cell is used only if its text begins with the exact prefix, including the colon. The number is the integer after the first `:`.
- The three cells are read in the order rows 1, 2, 3. The first one that fails to parse stops reading, and that setting and the ones after it keep their defaults.
- Values are not validated. The settings boxes clamp them to their ranges for display, but the unclamped values stay in the model until the user changes them.

### 4.6 `.xls`

`.xls` is accepted by the open dialog and by drag and drop, but the reader handles only `.xlsx`, so opening an `.xls` file fails with the load error (see section 9).

### 4.7 `.csv`

- One sheet, named after the file name **including the extension** (for example `data.csv`).
- Format:
  - comma-separated, UTF-8
  - the first line is the header
  - standard double-quote quoting
  - blank lines are skipped
  - every value is read as text, and empty cells are empty
- Entry names:
  - a blank header gets `Unnamed: {i}` (0-based position)
  - repeated header names get `.1`, `.2`, … appended to the later occurrences
  - (Replaced by sections 9.36–9.38: header names are kept exactly as read, blank and repeated ones included, so the file survives a round trip.)
- One card per data line, with rowSpan 1. A file with only a header line has **zero cards**.
- No pictures. Settings: width 136, height 180, 8 per row.
- No source workbook.
- An empty file fails with the load error.
- **Date-prefix display rule:** when a card shows a plain text value of at least 10 characters whose first 10 characters form a valid ISO date `YYYY-MM-DD`, the field shows that date in the short date format of the UI language and **discards the rest of the text** (see section 9).

---

## 5. Writing `.xlsx`

### 5.1 Workbook

1. **Start point.** Start from the source workbook if there is one; otherwise from a new empty workbook.
2. **Delete missing sheets.** Delete every worksheet whose name is not the name of an open sheet. This includes worksheets that were skipped at load because they were empty, and the original worksheet of a renamed sheet. A new workbook's default sheet is deleted too.
3. **Write each sheet** in tab order. If a worksheet of that name exists, write into it; otherwise append a new worksheet. Existing worksheets keep their position in the workbook; tab order does not reorder them.
4. **Save** the whole workbook to the chosen path.

Everything not mentioned below stays as it was: other cells, other sheets, styles, column widths, pictures outside the table, defined names and so on.

### 5.2 Per sheet: preparation

Let:

- `w` = picture width
- `h` = picture height
- `P` = pictures per row
- `N` = number of entries

1. **Cap pictures per row.** If `P × w > 1762`, then `P = max(1, floor(1762 / w))`.
2. **Metadata cells.** Write column 16384, rows 1–3: `ImageHeight:{h}`, `ExcelImagesPerRow:{P}` (the capped value) and `ImageWidth:{w}`. Hide column 16384.
3. **Table name.** If the table name is non-empty, write it to row H (or row 1 if H = 0), column tableNameColumn. If it is empty, nothing is written.
4. **Header.** If the headers are shown, write the header cells on row H+1, columns firstColumn … firstColumn+N−1: an empty string for unnamed entries, otherwise the name.
5. **Units per picture.** A picture is 2 units if it is wide (width ≥ height), otherwise 1. A picture that cannot be decoded counts as 1.
6. **Per-column maximum.** For each entry column c, `maxUnits[c]` = the largest total of units in any single card's value in that column (0 if none).
7. **Column picture capacity.** `colP[c] = maxUnits[c] > 0 ? max(1, min(maxUnits[c], P)) : P`.
8. **Column widths.** For each column with `maxUnits[c] > 0`:
   - target = `colP[c] × w` px
   - current = the column's width in px, using 8.43 if the column has no explicit width
   - if target > current, set the width to `pxToWidth(min(target, 1762))`

   Columns are only ever widened.

Width conversions (character-width units ↔ pixels):

```
pxToWidth(px)  = px <= 12 ? px / 12 : (px − 5) / 7
widthToPx(wd)  = wd <= 1  ? wd × 12 : wd × 7 + 5
```

### 5.3 Per card: rows

Cards are processed in order.

- `origRows` = the sum of rowSpan over all previous cards.
- `inserted` = the total number of rows inserted so far on this sheet.
- The card's first row: `R = dataStart + origRows + inserted`, where dataStart is H+2, or H+1 when the headers are hidden.

1. **Picture rows.** For each column with pictures in this card, `visualRows = ceil(totalUnits / colP[c])`. Take `V` = the largest of these, with a minimum of 1. Note `hasPictures` if any value in the card has a picture.
2. **Worksheet rows needed.** `need = max(rowSpan, ceil(V / 3))`. Each worksheet row holds up to **3 picture rows**.
3. **Inserting rows.** If `need > rowSpan`:
   - Insert `need − rowSpan` rows directly below row R.
   - Then, for every column from 1 to the sheet's last used column, merge rows R … R+need−1 vertically if either:
     - the column is a table column; or
     - any cell in rows R … R+need−1 of that column has a value.
4. **Re-merging.** Otherwise, if `rowSpan > 1`, merge every table column vertically over R … R+rowSpan−1.
5. **Row heights.** Only when hasPictures. For each `i` from 0 to need−1:
   - `k = min(3, V − 3i)`
   - if k > 0, set the height of row R+i to `(k × h + (k == 3 ? 6 : 0)) × 0.75` points

   Rows of cards without pictures keep their height.
6. **Unmerging.** For each merged range that contains row R:
   - Record, for each column it covers, its height in rows; per column, keep the largest.
   - If the range **starts** at R and either its height ≠ need or it covers more than one column, unmerge it.

### 5.4 Per card, per entry: pictures

For entry column c (worksheet column `col = firstColumn + c`):

1. **Remove old pictures.** Remove every picture anchored in column col whose anchor row lies in R … R+M−1. M = max(need, the recorded merge height for this column, or 1 if none).
2. **Place the value's pictures in order:**
   ```
   visualRow = 0; units = 0
   for each picture:
       wide = width >= height
       (tw, th, u) = wide ? (2w, h, 2) : (w, h, 1)
       if units + u > colP[c]: visualRow += 1; units = 0
       anchor cell   = (row R + floor(visualRow / 3), column col)
       offset x (px) = units × w
       offset y (px) = (visualRow mod 3) × h
       display size  = tw × th px
       units += u
   ```
3. **Crop.** Each picture is cropped to the target aspect `tw / th` by centre crop, expressed as a source rectangle in units of 1/100000 of the image size, truncated to integers:
   - **Too wide** (aspect > target): `newW = H_img × target`; `crop = W_img − newW`; `l = crop/2`; `r = crop − l`. Then `left = trunc(l / W_img × 100000)` and `right = trunc(r / W_img × 100000)`.
   - **Too tall** (aspect < target): `newH = W_img / target`; `crop = H_img − newH`; `t = crop/2`; `b = crop − t`. Then `top` and `bottom` are computed the same way against `H_img`.
   - **Same aspect:** no source rectangle.

   The original image bytes are embedded unchanged.
4. **Picture object.**
   - One-cell anchor at the anchor cell, with offsets in EMU (`px × 9525`) and size `tw × th` px in EMU.
   - Rectangle geometry, no outline, compressed state "print".
   - Numbered per sheet from 1: id = n, name `Image {n}`, description `Picture`.
5. **Undecodable pictures** are skipped silently: they are not written (see section 9).

### 5.5 Per card, per entry: cell styles

Write to cell (R, col). A cell that is the hidden part of a merge (not its top-left cell) is skipped for alignment and text.

**For styled and picture values:**

- **Fill:**
  - if fill is set: solid fill with start and end colour `FF` + RRGGBB, keeping the other fill properties;
  - if fill is none: remove the fill.
- **Text colour:**
  - if set: font colour `FF` + RRGGBB, keeping the other font properties;
  - if missing: replace the font with a default font.
- **Alignment, when the value carries alignment data:**
  - Start from the cell's current alignment.
  - Override horizontal and vertical when they are set, and wrapText when it is true or false.
  - Then, if the value has pictures or its text contains a newline, force wrapText = true, and vertical = top if no vertical alignment is set.
- **Alignment, when the value carries no alignment data:** if it has pictures or multi-line text, set the alignment to wrap + top. Otherwise leave it alone.
- **Font (when there is font data):**
  - Start from the cell's current font.
  - Set the name and size when present (non-empty, non-zero).
  - Set bold, italic, underline and strike when they are not unset.

**For plain values:** if the text contains a newline, set wrapText = true on the current alignment, and vertical = top if unset. Otherwise leave the cell untouched.

### 5.6 Per card, per entry: text

If the value's text is non-empty, write it to the cell **as a string**. An empty text writes nothing, so the old cell content stays (see section 9). Text beginning with `=` ends up stored as a formula.

### 5.7 Print area

After all cards:

- columns firstColumn … firstColumn+N−1
- first row = H+1, or the table name row if that is above it
- last row = dataStart + (sum of all rowSpans) − 1, but at least the first row

Rows inserted during the save are not counted in the last row (see section 9).

### 5.8 Loading summary per format

| Ext | Read | Write |
|---|---|---|
| `.xlsx` | Full (section 4) | Full (section 5) |
| `.xls` | Fails (section 4.6) | Not offered |
| `.csv` | Section 4.7 | Section 5.9 |

### 5.9 CSV export

- Exports only the **first** sheet in tab order, not the selected one.
- **Header line:** the entry names, with `""` for unnamed entries. For sheets with hidden headers the header line is `0,1,2,…`.
- **One line per card:** each value's text. Pictures, colours and styles are dropped. Row spans and the table name are not written.
- Format: comma-separated, UTF-8 without BOM, CRLF line endings, quoting only where needed.
- After the export the whole document is marked unmodified.

---

## 6. Pictures

### 6.1 Adding pictures

Every value can hold pictures. New pictures are appended to the end of the list. There is no way to reorder pictures within a value, or to move them between values.

1. **`Add Attachment` button.**
   - Opens the open-file dialog: title `Upload Images`, filter `Images (*.png *.jpg *.jpeg *.bmp *.gif)`, multiple selection.
   - Each selected file's bytes are appended unchanged.
   - Files that cannot be read are skipped silently.
2. **Paste** (Ctrl+V in the value's text field or the picture strip, or `Paste` in the strip menu):
   - If the clipboard holds file paths, every path ending in `.png`, `.jpg`, `.jpeg`, `.bmp` or `.gif` (case-insensitive) is added. Other paths are ignored, and image data on the clipboard is then ignored too.
   - Otherwise, image data on the clipboard is encoded as PNG and added.
3. **Drop** onto the value's text field or picture strip:
   - Image data is encoded as PNG and added.
   - Otherwise, every dropped file with one of those extensions is added.

Adding pictures marks the document modified.

### 6.2 Picture strip

- The strip is visible only when the value has at least one picture.
- It is a horizontal scrolling row, fixed at 130 px high: horizontal scroll bar as needed, never a vertical one, 5 px inner margin, 10 px spacing, left-aligned and vertically centred.
- Each picture is a **100 × 100 thumbnail**: a centre crop of the picture to a square, scaled with high-quality resampling.
  - Centre crop: if the source is wider than the target aspect, `newW = floor(H × aspect)`, `left = floor((W − newW)/2)`; for a taller source the same applies vertically.
  - A picture that cannot be decoded shows the text `Error` in a 100 × 100 box.
- **Remove button.** A 20 × 20 round "×" button overlaps the thumbnail's top-right corner: white text on red at 150/255 opacity, fully opaque red when pointed at, 16 px bold. It removes that picture immediately.
- **Remove All Attachments.** Under the strip, a `Remove All Attachments` button in the danger style removes every picture of the value immediately, without confirmation.
- **Interactions:**
  - Left-click a thumbnail to open the preview.
  - Right-click a thumbnail for `Copy selected photo`.
  - Right-click the strip background for `Copy all photos` and `Paste`.
- The strip can take keyboard focus, for Ctrl+V.
- The width, height and per-row settings have **no effect** on thumbnails.

### 6.3 Preview

- A modal dialog titled `Full Image Preview` shows the picture at its natural size.
- If the picture is larger than 80% of the primary screen's width or height, it is scaled down, keeping its aspect ratio, to fit within that box.
- The dialog has no buttons; close it with the window close button.

### 6.4 Copy

- **Copy selected photo:** puts the picture on the clipboard as image data.
- **Copy all photos:**
  - Writes each picture's bytes to a new temporary file with the extension `.png`, whatever the actual format.
  - Puts the list of those files on the clipboard.
  - The temporary files are never deleted.

### 6.5 Display sizes in the saved file

These follow section 5.4:

- a tall picture is `w × h`
- a wide picture is `2w × h`
- both are centre-cropped to that frame using a source rectangle, so the full original stays in the file

---

## 7. Appearance, languages and settings

### 7.1 Theme

The app follows the system light/dark setting at startup, and switches live when the system setting changes.

| Role | Light | Dark |
|---|---|---|
| Window and dialog background | #ffffff | #1e1e1e |
| Card background, selected tab | #f0f0f0 | #2d2d2d |
| Text | #222222 | #e0e0e0 |
| Disabled button text | #aaaaaa | #888888 |
| Borders | #cccccc | #444444 |
| Text field background | #ffffff | #3c3c3c |
| Text field text | #000000 | #eeeeee |
| Button and tab background | #eeeeee | #444444 |
| Button and tab hover | #dddddd | #555555 |
| Danger background (Remove All, tab close hover) | #ffcccc | #662222 |
| Danger text | #000000 | #ffffff |
| Card list (scroll area) background | #f9f9f9 | #252525 |
| List backgrounds in dialogs | text field background | text field background |

Other styling:

- **Buttons:** 1 px border, 5 px top and bottom padding, 15 px left and right, 3 px corner radius. Small buttons (`↑`, `↓`, the aspect lock) have 2 px left and right padding.
- **Danger buttons:** bold.
- **Text fields in cards:** 5 px padding, 1 px border, 3 px corner radius.
- **"+" add-sheet button:** transparent, borderless, 18 px bold. When pointed at, it shows a round hover background with a 14 px radius.
- **Pointer:** a pointing-hand pointer over every button, checkbox, drop-down and picture thumbnail.

### 7.2 Languages

- **Supported languages:** English (`en`) and Russian (`ru`).
- **Startup language:**
  1. the saved setting, if it is `en` or `ru`;
  2. otherwise the two-letter system UI language, if supported;
  3. otherwise `en`.
- **Switching** with the selector:
  - saves the choice;
  - immediately re-labels the top bar, the "+" tooltip, every sheet's settings bar, every card and every picture strip (the strips rebuild their thumbnails);
  - also sets the date format used for later loads.

  Dates that are already loaded keep their format. The window title, the `(Copy)` suffix and the `px` suffix are never translated.
- **Selector item text** is the language name translated into the current language.
- **Standard buttons** (Yes/No/OK/Cancel in message boxes and text-input dialogs), the colour dialog and the text-field context menu are not covered by the app's string table.

Every UI string, with its Russian translation:

| English | Russian |
|---|---|
| English | Английский |
| Russian | Русский |
| Upload Spreadsheet | Загрузить таблицу |
| Save Spreadsheet | Сохранить таблицу |
| Add Sheet | Добавить лист |
| Remove Sheet *(defined, unused)* | Удалить лист |
| Unsaved Changes | Несохраненные изменения |
| Loading a new spreadsheet will lose any unsaved changes. Continue? | При загрузке новой таблицы все несохраненные изменения будут потеряны. Продолжить? |
| Any unsaved changes will be lost. Are you sure you want to exit? | Все несохраненные изменения будут потеряны. Вы уверены, что хотите выйти? |
| Open Spreadsheet | Открыть таблицу |
| Spreadsheets (*.xlsx *.xls *.csv) | Таблицы (*.xlsx *.xls *.csv) |
| Excel (*.xlsx) / CSV (*.csv) | Excel (*.xlsx) / CSV (*.csv) |
| Sheet Name: | Имя листа: |
| New Sheet Name: | Новое имя листа: |
| Warning | Внимание |
| Sheet '{name}' already exists. | Лист '{name}' уже существует. |
| Sheet '{new_name}' already exists. | Лист '{new_name}' уже существует. |
| Confirm Remove | Подтверждение удаления |
| Are you sure you want to remove sheet '{name}'? | Вы уверены, что хотите удалить лист '{name}'? |
| Rename Sheet | Переименовать лист |
| Error | Ошибка |
| Could not load spreadsheet: {error} | Не удалось загрузить таблицу: {error} |
| Could not save spreadsheet: {error} | Не удалось сохранить таблицу: {error} |
| Success | Успех |
| Spreadsheet saved successfully! | Таблица успешно сохранена! |
| Table Name: | Имя таблицы: |
| No table name | Нет имени таблицы |
| Add Card | Добавить карточку |
| Remove Card | Удалить карточку |
| Add Entry | Добавить поле |
| Remove Entry | Удалить поле |
| Reorder Entries | Порядок полей |
| Image Width:␠ | Ширина изображений:␠ |
| Image Height:␠ | Высота изображений:␠ |
| Images per row:␠ | Изображений в ряду:␠ |
| Lock Aspect Ratio | Сохранять пропорции |
| Remove Entries - {name} | Удалить поля - {name} |
| Check entries to remove: | Выберите поля для удаления: |
| Reorder Entries - {name} | Порядок полей - {name} |
| Drag and drop to reorder entries: | Перетащите поля для изменения порядка: |
| Cannot remove the last card. | Нельзя удалить последнюю карточку. |
| Unnamed | Без названия |
| Add Attachment | Добавить вложение |
| Remove All Attachments | Удалить все вложения |
| Upload Images | Загрузить изображения |
| Full Image Preview | Предпросмотр изображения |
| Images (*.png *.jpg *.jpeg *.bmp *.gif) | Изображения (*.png *.jpg *.jpeg *.bmp *.gif) |
| OK | ОК |
| Cancel | Отмена |
| Background Color | Цвет фона |
| Text Color | Цвет текста |
| Pick Color... | Выбрать цвет... |
| No Fill | Без заливки |
| Duplicate | Дублировать |
| Rename | Переименовать |
| Remove | Удалить |
| New card above | Новая карточка выше |
| New card below | Новая карточка ниже |
| Copy selected photo | Копировать выбранное фото |
| Copy all photos | Копировать все фото |
| Paste | Вставить |

(␠ marks a trailing space: these three strings are prefixes placed directly before the number.)

Untranslated literal labels:

| Label | Where |
|---|---|
| `+` | add-sheet button |
| `🔗` | aspect lock |
| `↑`, `↓` | card move buttons |
| `×` | picture remove button |
| `px` | size boxes |
| `Spreadsheet Editor` | window title |
| ` (Copy)` / ` (Copy {n})` | duplicated sheet names |

### 7.3 Persisted settings

There is exactly one, stored per user:

| Setting | Key | Values | Default |
|---|---|---|---|
| UI language | `current_language` (group `Language`, application `SpreadsheetEditor`) | `en`, `ru` | System language if supported, else `en` |

Window size and position, recent files, last folder and picture settings are not persisted. Picture settings travel inside each saved `.xlsx` (section 4.5).

---

## 8. Change tracking, undo and performance

### 8.1 Modified flag

**Set to modified by:**

- adding, duplicating, renaming or removing a sheet (removing has an exception; see the next list)
- any card action: add, insert, duplicate, remove, move
- any entry action: add, rename, remove, reorder
- any text edit
- any colour change
- adding or removing a picture
- editing the table name
- changing the picture width, height or per-row setting

**Set to unmodified by:**

- a successful load
- a successful save (`.xlsx` or CSV)
- removing the last remaining sheet

**Not tracked:**

- reordering tabs
- toggling the aspect lock
- changing the language

**Shown:** as ` *` after the window title.

**Prompts:**

- the load confirmation (from the button, a drop or a paste)
- the exit confirmation

Both appear only when the document is modified and at least one sheet is open (section 2.12). Neither offers to save.

### 8.2 Undo and redo

The app has no undo or redo. Text fields keep their own text undo.

### 8.3 Performance

- **Building cards:** all cards of all sheets are built when a file loads. There is no lazy loading or virtualisation, and loading and saving run on the UI thread without progress feedback.
- **Incremental updates:** later structural changes are applied incrementally, and existing cards are never rebuilt:
  - adding, inserting or duplicating a card creates one card and renumbers the following cards;
  - removing a card deletes one card and renumbers the rest;
  - moving a card relocates one card;
  - adding, removing or reordering entries edits the rows inside each existing card;
  - renaming an entry updates only the name fields of the other cards.
- **Thumbnails:**
  - a value's thumbnails are decoded and rebuilt when one of its pictures is added or removed;
  - all thumbnails are rebuilt on a language switch.
- **Writing `.xlsx`:**
  - each picture is decoded to read its pixel size, and it is decoded again for each of the three layout passes;
  - when rows are inserted, the merge pass checks every column from 1 to the last used column. Because the metadata cells live in column 16384, that means up to 16,384 columns per card that needs extra rows.

---

## 9. Open questions and deliberate changes

Each item states the intended behaviour (**Intended**), the owner's answer (**Decided**, with its date) or asks the owner (**Q**). Items marked Intended or Decided are part of the build and replace what sections 1–8 say. Items marked Q keep the behaviour in sections 1–8 until the owner answers.

### File reading

1. **Metadata cells break detection when a saved file is reopened.**
   - The cells in column 16384, rows 1–3, count as occupied cells. When the table does not start on row 1, the row block starts at row 1, so the header is taken from an empty row 1. If the table starts more than 10 rows below row 3, the "table" becomes column 16384 alone.
   - **Intended:** ignore column 16384, and any cell holding a metadata string, during table detection.
2. **Dates.** Currently they become text and the time part is lost. **Decided (owner, 2026-11-03):** dates stay real date values.
   - A date cell is shown in the Windows short date format of the UI language, with the time when the cell has a time part. Time-only values are shown as a time.
   - A date value the user does not edit is written back as it was read: same serial value, same number format.
   - Edited text that parses as a date (or date and time) in the UI language is saved as a real date value and keeps the cell's number format (a cell without a date format gets the default short date format). Text that does not parse is saved as text.
3. **Numbers, booleans and formulas.** Every value is written back as a string, so `3` becomes the text `"3"` and `TRUE` becomes `"True"`. Formulas are shown as their source text and survive only because text starting with `=` is stored as a formula. **Decided (owner, 2026-11-03):**
   - Values the user does not edit keep their original type and are written back as read.
   - Edited text that is a number in the UI language's number format is saved as a number; `TRUE` and `FALSE` are saved as booleans; anything else is saved as text.
   - Formulas are shown and edited as their source text, beginning with `=`, and saved as formulas. The app does not compute them; the saved file makes the spreadsheet program recalculate them when it opens the file (how is the writer task's to propose, with a source).
   - Details settled with the cell-values task (manager, 2026-11-03, within the owner's answer): booleans show as `TRUE` / `FALSE` in every language and are typed in either case; a number is shown in its shortest exact form with the language's decimal separator; typed text becomes a number only when it is plainly one (no group separators, exponent, currency or percent, no leading zeros such as `007`, at most 15 significant digits), a time alone becomes a time, and anything else stays text; in a cell formatted as Text, whatever is typed stays text, as in Excel.
4. **ISO-date prefix truncation.** Text like `2024-01-15 meeting notes` from CSV is displayed (and later saved) as just the date. **Intended:** never alter text that is not a real date value.
5. **Indexed (legacy palette) colours** are ignored: the fill becomes none and the text becomes black. **Intended:** resolve them with the standard 64-colour indexed palette.
6. **Automatic text colour becomes explicit black** on every saved cell, and there is no way to set a text colour back to automatic. **Decided (owner, 2026-11-03):** "automatic" is a distinct text-colour state. A cell without a font colour reads as automatic and is never written as explicit black; the text-colour flyout has an `Automatic` item, as the fill flyout has `No fill` (section 10.11).
7. **Grouping uses merges in any column.** Values are read only from a card's first row, and on save every table column is merged over the card's rows. As a result, values in the second and later rows of a multi-row card that were not merged are invisible in the app and destroyed on save. **Decided (owner, 2026-11-03):**
   - Only merged ranges that lie inside the table columns group rows into a card; merges outside the table columns are ignored by grouping (and kept as they are).
   - Values in a card's second and later rows are not shown, and are kept in the file unchanged: a save never merges a column over a cell that holds a value in those rows (see item 24). Removing the card removes its rows, and those values with them.
8. **Empty rows inside the table become empty cards.** Q: keep them, or skip fully empty rows?
9. **Header row containing a number hides all headers**, and booleans count as numbers. **Decided (owner, 2026-11-03):** the headers are hidden only when **every** non-blank cell of the header row is a number. Booleans, dates and text do not count as numbers, so a row such as `Name | 2024 | Price` stays a header.
10. **Empty sheets are skipped at load and then deleted on save.** **Intended:** keep sheets the app does not show, untouched, in the saved workbook. Q: should they appear as tabs instead? **Decided (owner, 2026-11-03):** an empty worksheet (no occupied cell outside the layout metadata column) opens as a tab with a new sheet's defaults (section 2.3), as hidden worksheets do, so a document saved before anything was typed reopens as it was left. Saved without edits, its worksheet is written back unchanged. Chart sheets and other non-grid sheets stay skipped and kept (item 12).
11. **`.xls`** is offered but always fails. Q: support `.xls` reading (saving to `.xlsx`), or remove it from the filter and the drop targets? **Decided (manager, 2026-11-03; the owner may overrule):** remove it from the open filter and the drop targets, as Windows apps list only what they open. A legacy or encrypted file renamed to `.xlsx` still fails with the reader's clear message.
12. **Chart sheets** or other non-grid sheets may make the whole load fail. **Intended:** skip and preserve them.
    - **Hidden worksheets** (not covered above). **Decided (owner, 2026-11-03):** read them like every other worksheet (section 4.1) and show them as tabs; their hidden state is written back unchanged.
13. **Failed load clears the open document**, because tabs are removed before parsing. **Intended:** parse first, then replace the document only on success.
14. **CSV with only a header gives a sheet with zero cards.** **Decided (owner, 2026-11-03):** such a file opens with one empty card, matching the "at least one card" rule. Saving it without typing anything writes the header line alone.
15. **CSV encoding and delimiter** are fixed at UTF-8 and comma. **Decided (owner, 2026-11-03):**
   - The separator is detected from the header line: comma, semicolon or tab (the exact rule is the CSV task's to propose).
   - Encodings: UTF-8 with or without a BOM, and UTF-16 with a BOM. No legacy code pages.
   - Saving a sheet read from a `.csv` keeps that file's separator and encoding; other sheets are exported as section 5.9 says.

### File writing

16. **Clearing text does not clear the cell.** An empty value writes nothing, so the old content reappears in the file. **Intended:** an empty value clears the cell's value.
17. **Clearing the table name leaves the old name in the file.** **Intended:** clear that cell.
18. **Adding a table name to a table that had none** writes it into row 1 or row H. It is then overwritten by the header, or overwrites the header or data. **Decided (owner, 2026-11-03):** giving a name to a table read without a name row inserts a worksheet row directly above the header on save, shifting everything below it as in item 20, and writes the name there in column tableNameColumn. Clearing the name later clears that cell (item 17); the row stays. On reopening, the name is found by section 4.1 step 4, which needs more than one non-blank header cell; a table with a single non-blank header cell is not seen as named (unchanged).
19. **Removed cards and entries leave stale data.** Rows and columns beyond the new table size keep their old values, pictures and merges. **Intended:** clear the table region of the previous layout (values, pictures, merges) before writing.
20. **Adding cards overwrites whatever lies below the table**, and inserted rows shift only cell values. **Intended:** whenever rows are inserted, shift everything below them, including values, styles, merges, pictures, row heights and formula references. **Decided (owner, 2026-11-03):** the table region is rewritten with worksheet rows inserted or deleted, so that content below the table keeps its position relative to the table's end. Adding cards (or rows for pictures) inserts rows; removing cards, or a card needing fewer rows, deletes rows. Inserting and deleting both shift everything below: values, styles, merges, pictures, row heights, formula references, defined names and the print area. Whole worksheet rows move, so content beside the table in those rows moves with them.
    - Moved cards (owner, 2026-11-03, replacing the manager's earlier positional choice): references to a moved card's cells follow the card, as after cut and insert in Excel: a formula that pointed at Cherry's quantity points at the row Cherry is in after the save.
    - Cards added at the end (owner, 2026-11-03): a range that reaches the table's last row (for example a total `=SUM(B2:B6)` over a whole table column) grows over the rows of cards added at the table's end, as an Excel Table's references do. Rows inserted inside a range grow it as in Excel; a range that ends above the table's last row does not grow.
    - Cards added at the top count too: a range covering every card starts at the table's first data row after the save (manager, 2026-11-03, within the owner's answer; the owner may change it).
21. **Saving twice inserts rows twice.** The in-memory source workbook is modified by each save while the card row spans stay as loaded. **Intended:** saving is repeatable. Either each save starts from an unmodified copy of the source, or the row spans and source are updated after a successful save, so that two saves in a row give identical files.
22. **Row spans are attached to positions, not cards.** Inserting, removing or moving cards gives cards the wrong span, and reordering entries resets all spans to 1. **Intended:** a span belongs to its card and moves with it. New cards have span 1, and a duplicated card copies its source's span. Entry operations never change spans. A duplicated card copies its source rows' styles but not the values in its second and later rows, which are not in the model (manager, 2026-11-03, within the owner's answer; the owner may change it)).
23. **Horizontal merges in a card's first row are unmerged on save**, and so are vertical merges of a different height. Q: preserve horizontal merges inside the table?
24. **Multi-row cards merge every table column on save**, even columns that were not merged. Related to item 7. Q: confirm that "one card = every column merged" is the intended output. Partly answered by item 7: a column is never merged over a value in a card's second or later rows.
25. **Picture row count is under-estimated.** `ceil(units / capacity)` ignores wrap waste: with capacity 2, the pictures tall, wide, tall need 3 rows but get 2. Also, a wide picture in a column of capacity 1 always wraps first, which leaves an empty picture row. **Intended:**
    - count rows by running the same placement loop used to position pictures;
    - a picture is always placed on the current row when that row is empty, even if it is wider than the capacity.
26. **3 picture rows per worksheet row** fits the worksheet's maximum row height (409.5 pt = 546 px) only when `h ≤ 180`: 3 × 180 + 6 = 546. Larger heights produce rows that the spreadsheet clips. **Decided (owner, 2026-11-03):** `rowsPerWorksheetRow = max(1, floor(540 / h))` (10 at `h` = 50, 3 at 180, 1 from 271 to 500). It replaces the 3 in sections 5.3 (steps 2 and 5) and 5.4 (step 2): `need = max(rowSpan, ceil(V / r))`, `k = min(r, V − r × i)`, a full worksheet row (`k = r`) adds 6 px as before, and the anchor row and y offset use `floor(visualRow / r)` and `(visualRow mod r) × h`. Since `r × h ≤ 540`, every height in the box's range (50–500) fits.
27. **Print area ignores inserted rows.** **Intended:** the last print row includes every row the table occupies after insertion. Further (manager, 2026-11-03, within the owner's answer; the owner may change it): a print area the sheet already has is kept and shifted with inserted and deleted rows like any other reference; section 5.7's print area is written only when the sheet has none, so a user's own print settings survive.
28. **Undecodable pictures are silently dropped on save.** **Intended:** keep the bytes. Write them as a 1-unit picture at their stored size, or warn before saving.
29. **The merge pass scans every column up to the last used column** (16384, because of the metadata). **Intended:** consider only the table columns plus columns that actually have values in the affected rows. Never merge column 16384.
30. **Sheet rename and duplicate on `.xlsx`.** A renamed sheet is deleted and recreated, losing everything outside the table: other cells, column widths, page setup and pictures. A duplicated sheet starts blank apart from the table. **Intended:** rename the worksheet in place, and duplicate by copying the whole worksheet before writing the table.
31. **Tab order does not change sheet order in the saved workbook**, and dragging tabs does not mark the document modified. **Intended:** saved sheet order follows tab order, and reordering tabs marks the document modified.
32. **Duplicate sheet uses the table name as loaded**, not the edited one. **Intended:** use the current field text.
33. **Removing every entry is allowed**, and saving then fails. **Intended:** the Remove Entries dialog cannot remove the last entry. Its OK button is disabled when every entry is checked.
34. **CSV export writes the first sheet, not the selected one**, writes `0,1,2,…` as the header for sheets with hidden headers, and marks the whole document saved. **Decided (owner, 2026-11-03):**
    - export the selected sheet;
    - omit the header line for a sheet with hidden headers;
    - the document stays modified when it has other sheets that were not written.
35. **The `.csv` check on save is case-sensitive** (`X.CSV` gets `.xlsx` content). **Intended:** decide the format by case-insensitive extension, or by the chosen filter when there is no extension.

### Editing

36. **Entries are matched by name.** Duplicate or empty names make Remove Entries and Reorder Entries act on the wrong entries, or on several at once. **Intended:** entries have a stable internal identity, and the dialogs work by identity and position.
37. **Hidden-header sheets** list their entries as `Unnamed`, `1`, `2`, … in the dialogs, because entry 0 counts as unnamed. **Intended:** list them as `Entry 1`, `Entry 2`, … (needs a translation, Russian `Поле 1`, …). Q: should users be able to name entries on such sheets, which would turn headers on?
38. **Typing an entry name that begins with `Unnamed:`** makes it count as unnamed: it is written as an empty header and shown empty next time. **Intended:** only empty names are unnamed. Placeholder names are internal and never collide with user text.
39. **Sheet name rules.** Names are checked only for exact duplicates. **Intended:**
    - enforce the worksheet name rules: 1–31 characters, none of `: \ / ? * [ ]` (nor U+0000 or U+0003), no leading or trailing apostrophe, not `History` (reserved by Excel), and not already used ignoring case (Microsoft's rules as quoted in `SheetNames`, read 2026-11-03);
    - show the reason inline in the dialog;
    - for duplicated names, shorten the base name so that ` (Copy n)` fits.
40. **Double-clicking the empty part of the tab strip** opens Rename with an empty name. If confirmed, nothing is renamed but the document is marked modified. **Intended:** a double-click on empty strip space does nothing. Q: should it add a sheet instead?
41. **Removing the last sheet marks the document unmodified**, so the app can close without a prompt after the user deleted everything from a loaded file. Q: is that intended?
42. **`↑` on the first card and `↓` on the last card** do nothing. **Intended:** disable them.
43. **The fill and text colour are not shown in the text fields.** Q: preview the fill, text colour, bold, italic and alignment in the field?
44. **Picture settings do not affect in-app thumbnails** (always 100 × 100 squares). Q: show thumbnails at the configured aspect, with wide pictures double width, so the card previews the file layout?
45. **Pictures cannot be reordered** within a value, or moved between values. Q: add drag-to-reorder and move?
46. **`as_text` is never read.** Q: what should it mean? One possibility: "write this value's pictures as text, such as file names, rather than embedding them". Otherwise drop it.
47. **"Picture entries"** only affect what an empty value looks like in a new card. Q: remove the concept?
48. **Paste with file paths and image data on the clipboard** uses only the paths. If none of them is a picture, nothing is added. **Intended:** fall back to the image data.
49. **Copy all photos** writes every picture as `.png` whatever its real format, and never deletes the temporary files. **Intended:**
    - use the real format's extension;
    - put the files in an app-specific temporary folder;
    - clean up that folder on exit and on the next start.
50. **Metadata values are not validated.** **Intended:** clamp values read from the file to the box ranges, with pictures per row clamped to the width-dependent maximum.
51. **Aspect lock** is per session and its ratio is not recaptured after clamping. Q: persist the lock in the metadata cells (for example row 4, `AspectLocked:1`)?
52. **Unused string:** `Remove Sheet` has a translation but no UI. Q: add a `Remove Sheet` command, for example in a menu, or drop the string?
53. **Performance.** **Intended:**
    - virtualise the card list;
    - decode thumbnails off the UI thread and cache them;
    - read and write files asynchronously, with a progress indicator and the UI disabled during the operation;
    - decode each picture's size once per save.

---

## 10. Windows conventions

**Decided (owner, 2026-11-03): follow Windows' design and functionality guidelines** (ADR 0001). Each item gives the Windows 11 alternative, and the app builds that alternative instead of what sections 1–8 say; every "Q" below is answered yes. Where an item offers a choice (12, 15, 16, 23), the task that builds it proposes the native option in its plan and the manager confirms it with the owner.

1. **Tab close buttons and the floating "+".** These are drawn by hand. Use the standard tab view, which has a built-in add-tab button and close buttons, middle-click to close, Ctrl+T, Ctrl+W and Ctrl+Tab, and inline rename (double-click the header to edit in place) instead of a rename dialog. Q: switch?
2. **Pointing-hand pointer on buttons, drop-downs, checkboxes and thumbnails.** Windows uses the normal arrow for controls, and the hand only for hyperlinks. Q: use the arrow?
3. **Home-made dialogs** (Remove Entries, Reorder Entries, sheet name input, preview). Q: use the standard content dialog, with the primary button on the left, Cancel on the right, Enter and Esc handling, and system-localised button text?
4. **Yes/No unsaved-changes prompts with no Save option.** The Windows pattern is "Do you want to save changes to {file}?" with **Save**, **Don't save** and **Cancel**. Q: adopt it for exit, open and drop?
5. **Every save asks for a path.** Windows apps have **Save** (Ctrl+S, to the current file, asking for a path only the first time) and **Save as** (Ctrl+Shift+S). They suggest the current file name and folder, and remember the last folder. Q: adopt this, along with **Open** (Ctrl+O) and **New** (Ctrl+N)?
6. **"Upload Spreadsheet"** is web wording. Windows uses **Open**. Likewise, **Add Attachment** could be **Add pictures**. Q: rename?
7. **A success message box after every save.** Windows apps save silently: the title marker clears, and at most a brief non-blocking notice appears. Q: remove the box?
8. **Window title.** The title shows only `Spreadsheet Editor *`. Windows convention is `{file name} - Spreadsheet Editor`, with an unsaved marker. Q: show the file name?
9. **Language switch inside the app.** Windows apps follow the Windows display language, with an optional override in a Settings page. Language names are shown in their own language (`English`, `Русский`), never translated. Q: drop the top-bar selector and follow Windows, with an override in Settings?
10. **Standard button text is not localised** in Russian (Yes/No/OK/Cancel). With native dialogs these come from the system. Q: confirm native dialogs.
11. **Colour dialog.** Q: replace it with the standard colour picker in a flyout attached to the swatch, with a row of recent and theme colours, plus `No Fill` and an optional `Automatic` for text?
12. **Enter moves to the next field and Shift+Enter inserts a newline** in multi-line fields, while Tab inserts a tab character. In Windows multi-line text boxes, Enter inserts a newline and Tab moves focus. Q: keep the form-filling behaviour (Enter = next field) or follow the platform? Separately, Tab should move focus in either case. Also, what should Enter do in the last field of the last card?
13. **Card reordering by ↑/↓ buttons only.** The Windows pattern is drag-to-reorder in a list, plus Alt+Up / Alt+Down. Q: add both and keep the buttons?
14. **Destructive actions without confirmation or undo** (Remove Card, Remove All Attachments, picture ×, Remove Entries). The Windows pattern is either confirmation or undo (Ctrl+Z). Q: add application-level undo/redo, which would let most confirmations go?
15. **Ctrl+V on the window opens a copied spreadsheet file.** This is unusual on Windows, and it clashes with paste expectations. Q: keep it? **Decided (manager, 2026-11-03; the owner may overrule):** no. Ctrl+V only pastes; files open with Open, drag and drop, or the recent files list.
16. **Picture preview is a fixed-size modal dialog** with no zoom and no close button. Q: use a separate resizable window, or a full-window overlay with zoom, Esc to close and arrow keys to step through the value's pictures?
17. **Raw error text in messages.** Q: show a plain explanation (file in use, access denied, unsupported format) with the technical details in an expander?
18. **Custom colours.** Hard-coded light and dark colours replace the system theme. Q: use the system theme resources (backdrop material, accent colour for the selected tab and primary buttons, high-contrast support)?
19. **Number boxes show their caption inside the box** (`Image Width: 136px`). The Windows pattern is a number box with a header label above or beside it, and units in the header. Q: switch?
20. **The aspect-lock toggle is an emoji 🔗.** Q: use the system "link" icon glyph on a toggle button, with an accessible name?
21. **Keyboard and accessibility.**
    - Context menus open only by mouse.
    - The colour swatches have no accessible names beyond tooltips.
    - The picture strip has no keyboard way to open, copy or remove a picture.

    Windows expects Shift+F10 and the Menu key for context menus, Delete on a focused picture, Enter to preview, and screen-reader names for every control. Q: implement all of these?
22. **Context menus without icons or shortcuts.** Windows 11 menus show icons and accelerators, such as Ctrl+C for copy and Delete for remove. Q: add them?
23. **Sheet removal always asks**, even for an empty sheet. Spreadsheet programs on Windows ask only when the sheet has content. Q: ask only for non-empty sheets, or replace the prompt with undo?
