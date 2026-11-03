# Spreadsheet Editor

A Windows app that edits a spreadsheet as a form: each row of a sheet's table is a card, each column an entry on the card, and image
columns hold pictures laid out inside their cells. It opens and saves `.xlsx` and `.csv`. Solo project. The product's behaviour is
specified in `docs/spec.md`; the plan is `docs/roadmap.md`.

**A native Windows app** (ADR 0001): WinUI 3 (Windows App SDK) in C# on .NET 10, over a portable core library (ADR 0002).

## Principles
- **All logic lives in the core** (`SpreadsheetEditor.Core`): the document model, reading and writing files, the image layout,
  editing operations and their undo, validation. The app renders the core's state and sends user actions; it holds only platform
  integration (window, pickers, clipboard, drag and drop, settings storage, dialogs).
- **Native behaviour by default.** The app looks and behaves as Windows 11's own apps do: Fluent controls and typography, Mica, the
  title bar, menus, keyboard accelerators and access keys, focus, accessibility (UI Automation), light, dark and high contrast. Where
  Windows has a behaviour, keep it. Departing from it needs a reason and the owner's agreement, recorded in a decision record.
- **Fast with large files.** Load and save off the UI thread, virtualise card lists, decode images at display size and lazily. Nothing
  that grows with the file runs on the UI thread.
- **Files survive a round trip.** Whatever the app does not edit (other sheets' content, styles, formulas, merged cells outside the
  table, print settings) is written back as it was read. `docs/spec.md` says what the app may change.

## Layout
```
src/SpreadsheetEditor.Core    the core: model, xlsx/csv I/O, image layout, editing and undo (net10.0, no Windows API)
src/SpreadsheetEditor.App     the WinUI 3 app (builds SpreadsheetEditor.exe): unpackaged, self-contained
tests/SpreadsheetEditor.Core.Tests   MSTest tests of the core
docs/                         spec.md, architecture.md, roadmap.md, decisions/ (ADRs), notes/
build.ps1                     publish into out/app; -Smoke starts the app once and checks it exits cleanly
tools/icon/make_icon.py       draws the app icon (development only; Python with Pillow)
```
A new project copies the shape of `SpreadsheetEditor.Core` and is added to `SpreadsheetEditor.slnx`. Test fixtures (small `.xlsx`,
`.csv` and image files) live in `tests/SpreadsheetEditor.Core.Tests/Fixtures/`; prefer building workbooks in code over checking in
binary files.

## Commands
```
dotnet test tests/SpreadsheetEditor.Core.Tests                     # core build and tests; anywhere .NET 10 runs
dotnet format whitespace --verify-no-changes <project>             # formatting, per changed project
dotnet build src/SpreadsheetEditor.App -c Release -r win-x64 -p:Platform=x64   # Windows only
pwsh ./build.ps1 -Smoke                                            # Windows only: publish and smoke-run the app
```
CI (`.github/workflows/ci.yml`) runs the core on Linux and Windows and publishes, smoke-runs and uploads the app on Windows (x64, and
builds ARM64). The app (WinUI) cannot be built in a Linux cloud container: say so in the PR and let CI build it. In a cloud container,
check `dotnet --version` first; if .NET 10 is missing, you may install the SDK into the container only, outside the repo, with
Microsoft's script and the repo's `global.json` (`dotnet-install.sh --jsonfile global.json --install-dir ~/.dotnet`; owner,
2026-11-03), then run the commands above. If the install fails or the session refuses it, do not retry or work around it: CI builds
and tests the core, and the PR says so.

## Conventions
- C# latest, nullable enabled, warnings are errors (`Directory.Build.props`). File-scoped namespaces, one public type per file.
- The core has no reference to WinUI, WinRT or any Windows API; `IsAotCompatible` stays true (no reflection-based serialisation).
- The app uses MVVM without a framework: view models in `ViewModels/`, `INotifyPropertyChanged` by hand, `x:Bind` over `Binding`.
- Every user-visible string is a resource (`Strings/<lang>/Resources.resw`), English and Russian; no string literals in XAML or code
  for UI text.
- Errors are exceptions with a message the app can show; the app shows them in a `ContentDialog` or `InfoBar`, never crashes on a bad
  file.
- Tests: unit tests next to the feature in the test project, named `<Type>Tests`. Round-trip tests write a workbook, read it back and
  compare. A test that cannot run is skipped visibly (`[Ignore("reason")]`), never passes by returning early.
- **Dependencies:** never add a package, tool or GitHub Action, or bump a version (the .NET SDK included), without the owner's approval.
  Every package is declared once in `Directory.Packages.props` with an exact version (`[x.y.z]`); projects reference it without a
  version. Licences per ADR 0003. Actions are pinned by commit SHA with the tag in a comment.
  Standing exception: Microsoft's own Windows platform packages (the Windows SDK, the Windows App SDK and its component packages, the
  Windows Community Toolkit) are approved in advance. Pin exact versions, check their licences and dependencies under ADR 0003, and
  report what was added in the PR.
- **Do not assert API or platform facts without a source.** Quote the source and the date read; mark anything unchecked "not verified".
- **The repo is public.** No personal data, tokens or keys in code, tests, fixtures, docs or commits: no real names, addresses, paths from the owner's machine or file
  contents from the owner's spreadsheets. Fixtures are made up.
- Never add an about or licences screen unprompted: the owner says when.

## Git workflow
- `main` always builds and passes CI: PRs only, squash merges only, never force-push.
- One task = one branch = one PR. Keep PRs small. Never reuse a branch after its PR is merged.
- Branch names are `<type>/<topic>`. Type: `feat`, `fix`, `docs`, `refactor`, `test`, `build`, `chore`, or `spike` (an experiment whose
  PR carries only its decision record or notes). Topic: a few lowercase words joined by hyphens.
- Bring main into a branch with a merge commit, never a rebase.
- Commit messages: imperative, short subject, optional body explaining why. No `Co-Authored-By`, `Claude-Session` or "Generated with
  Claude Code" lines in commits or PR descriptions.
- Author and committer of every commit: `skeksik <personal@skeksik.com>` (in a fresh environment, `git config user.name skeksik` and
  `git config user.email personal@skeksik.com` before the first commit).
- Before a PR, run the commands above and report failures honestly. Fill in `.github/pull_request_template.md`, including what was not
  verified.
- Record architectural decisions as short ADRs in `docs/decisions/` (next free number).

## Working with Claude
- Non-trivial work: propose a plan first, then implement in small steps.
- `/manage` runs a cloud session as the manager that starts worker sessions, reviews their PRs and merges after the owner's rules
  (`.claude/commands/manage.md`).
