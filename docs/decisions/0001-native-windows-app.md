# 0001: A native Windows app

Status: accepted (2026-11-03)

## Context
Spreadsheet Editor edits a workbook's table as a list of cards, with pictures inside image columns. Workbooks reach hundreds of rows
and hundreds of pictures, so opening, scrolling and saving must stay fast, and the app must feel at home on Windows 11, where it is
used. A cross-platform UI toolkit draws its own controls and imitates the platform's behaviour by hand; it is slower to start, heavier
in memory and never quite matches Windows' keyboard, focus, accessibility and materials.

## Decision
- **WinUI 3 (Windows App SDK) in C#** on .NET 10, for Windows 10 1809 (17763) and later, x64 and ARM64. Unpackaged and self-contained
  first; an installer or MSIX later (docs/roadmap.md).
- **Native behaviour by default.** The app follows Windows' design guidance and its own apps: Fluent controls and typography, Mica, the
  title bar, MenuBar and CommandBar, keyboard accelerators and access keys, ContentDialog for confirmations, InfoBar for status, the
  system file pickers, the clipboard and drag and drop, light, dark and high contrast, UI Automation. Where Windows has a behaviour, we
  keep it. Departing from it needs a reason (Windows' behaviour is questionable for this app) and the owner's agreement, recorded in a
  decision record.
- **Behaviour comes from `docs/spec.md`**, which says what the app does. Where the spec and Windows conventions disagree (a custom
  control where Windows has one, a non-standard shortcut, a home-made dialog), Windows wins unless a decision record says otherwise.

## Consequences
- Accessibility comes from WinUI (UI Automation), not rebuilt by hand. Checks with Narrator, high contrast and keyboard only happen once
  per milestone (docs/roadmap.md).
- The app cannot be built in a Linux cloud container; CI builds it on Windows runners, and hands-on checks run on the owner's PC.
