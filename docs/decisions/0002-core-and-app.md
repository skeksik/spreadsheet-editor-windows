# 0002: A portable core library and a thin WinUI app

Status: accepted (2026-11-03)

## Context
The app's rules (how a sheet's table is found, how rows become cards, how pictures are laid out and cropped in cells, what a save
changes) are most of its value and most of its risk. They must be testable without a window and in a Linux cloud container, where
worker sessions run.

## Decision
- **`SpreadsheetEditor.Core`** (net10.0, no Windows API) holds the document model, reading and writing `.xlsx` and `.csv`, the image
  layout, editing operations with undo and redo, and every rule in `docs/spec.md`. It is AOT-compatible: no reflection-based
  serialisation.
- **`.xlsx` through the Open XML SDK** (`DocumentFormat.OpenXml`, Microsoft, MIT). The core edits the opened package in place, so
  everything it does not model (other sheets, styles, formulas, defined names, drawings it does not own) is kept as it was. No
  higher-level spreadsheet library: they rewrite whole workbooks and lose what they do not model.
- **`.csv` by our own reader and writer** (RFC 4180, UTF-8 with or without BOM, comma separator, `"` quoting).
- **Image facts in the core**: the core reads a picture's pixel size from its header (PNG, JPEG, GIF, BMP) without decoding it. The app
  decodes pictures (WinRT `BitmapImage`/`BitmapDecoder`) at the size it shows them.
- **`SpreadsheetEditor.App`** (WinUI 3) renders the core's document through view models and sends user actions to it. It holds only
  platform integration: the window, pickers, clipboard, drag and drop, settings storage, dialogs, resources.
- **Threads**: opening and saving run on the thread pool; the document is edited on the UI thread only; the core is not thread-safe and
  says so.

## Consequences
- Cloud workers can build and test the core with .NET 10 where the container has it; the app is built by CI.
- Every behaviour has a core test; the app's own tests are limited to the smoke run until a UI test harness is chosen (docs/roadmap.md).
