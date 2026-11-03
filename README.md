# Spreadsheet Editor

Edit a spreadsheet as a form. Each row of a sheet's table becomes a card, each column an entry on the card; image columns hold pictures
laid out inside their cells. Opens and saves `.xlsx` and `.csv`.

A native Windows 11 app: WinUI 3 on .NET 10, x64 and ARM64.

## Build
Requires Windows, the .NET 10 SDK and PowerShell 7.

```
dotnet test tests/SpreadsheetEditor.Core.Tests
pwsh ./build.ps1          # publishes into out/app
```

See `CLAUDE.md` for conventions, `docs/spec.md` for behaviour and `docs/roadmap.md` for the plan.
