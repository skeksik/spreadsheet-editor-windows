# Localisation

How the app picks its interface language, where its strings live, and how to add one (spec §7.2 as replaced by §10.9-10.10).

## Strings
- Every user-visible string is in `src/SpreadsheetEditor.App/Strings/en-US/Resources.resw` and `Strings/ru-RU/Resources.resw`.
  Every key is in both files.
- **XAML** sets strings with `x:Uid`. The key is `<Uid>.<Property>`, e.g. `MainWindow_ExitItem.Text`, and the Uid is
  `<View>_<Element>`.
- **Code** reads keys without a dot through `Services/Strings`: `Strings.Get("AppDisplayName")`, or
  `Strings.Format("WindowTitleWithFile", fileName)` for `{0}` placeholders. Placeholders are positional, so a translation can move
  them. `Window.Title` is set in code because `Window` is not a `DependencyObject`, so `x:Uid` cannot reach it (not verified).
- **Never translated:** the product name `Spreadsheet Editor` (`AppDisplayName`, `WindowTitleWithFile`, §7.2), and the language names
  `English` and `Русский`, which are shown in their own language (`LanguageName_English`, `LanguageName_Russian`, §10.9). These keys
  have the same text in both files.
- **Window title:** `Spreadsheet Editor` with no document open. With a document open it is `{file} - Spreadsheet Editor`
  (`WindowTitleWithFile`, §10.8); feat/app-document sets it.
- **Standard dialog buttons** come from the system when the dialogs are native (§10.10). A `ContentDialog`'s button texts are still
  the app's own strings.

## Language
`Core/Localisation/UiLanguage` holds the rule and is tested. `App/Services/LanguageService` applies it:
1. The stored override (`en` or `ru`) wins.
2. Otherwise (`system`, missing or unknown), the app uses Windows' display language when it is English or Russian, in any region.
3. Otherwise it uses English.

At startup, in the `App` constructor before `InitializeComponent`, the service sets
`Microsoft.Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride` to `en-US` or `ru-RU`. It also sets
`CultureInfo.DefaultThreadCurrentUICulture` and `CurrentUICulture`, and exposes the culture as `LanguageService.UiCulture`. Later tasks
pass that culture to the core's value formatting (`Values/CellText`) and to the reader and writer. `CurrentCulture`, Windows' regional
format, is left as it is.

**A changed override takes effect at the next start.** `SavePreference` stores the choice, and `RestartNeeded` tells the Settings page
(feat/app-settings) to say so. Sources, read 2026-11-03:
- Microsoft Learn, "ApplicationLanguages.PrimaryLanguageOverride Property (Microsoft.Windows.Globalization)", Windows App SDK 1.6-2.0:
  - "If this is a packaged app, the setting is persisted between app sessions. Otherwise, it is not persisted between app sessions and
    should be set each time the app is loaded."
  - "Set PrimaryLanguageOverride setting during app load, before any resource is loaded."
  - "for resources already loaded in the app UI, this change might not take effect immediately. To ensure the app responds to these
    changes, you may need to reload resources (this can vary with UI framework, and may require an app restart)."
  - "Resources loaded manually with ResourceLoader and XAML resources fetched with x:Uid" follow the override (the page's example).
  - "For unpackaged apps, languages available for the app to use are limited to languages for which your app has explicitly declared
    support." The page does not say how an unpackaged app declares that. We take it to be the languages in the app's PRI file (not
    verified). The smoke check below confirms that both are there.
- Microsoft Learn, "CultureInfo.CurrentUICulture Property", .NET 10, on the Windows display language that the app follows: "On systems
  that use the Windows operating system, the common language runtime calls the Windows GetUserDefaultUILanguage function to set the
  current UI culture. GetUserDefaultUILanguage returns the default UI culture set by the user." `LanguageService.WindowsLanguage` reads
  it before the service changes it.
- Microsoft Learn, "Manage resources MRT Core": "The resources.pri file at the root of each package is automatically loaded when the
  ResourceManager object is instantiated." Which file name an unpackaged app loads, `resources.pri` or `SpreadsheetEditor.pri`, is not
  verified. `build.ps1` accepts either, and the smoke check proves the strings load.
- Build tools: `Microsoft.Windows.SDK.BuildTools.MSIX` 1.7.251221100 (the version this build restores) adds `**/*.resw` as
  `PRIResource` items (`Common.props`). Its `DefaultLanguage` defaults to `en-US` (`MrtCore.PriGen.targets`), and the app's project
  states the same value explicitly.

## Settings file
`App/Services/SettingsStore` keeps the user's settings in `%LOCALAPPDATA%\SpreadsheetEditor\settings.json`:

```json
{
  "language": "system"
}
```

- `language` is `system`, `en` or `ru`. A missing file, missing key or unknown value means `system`.
- The JSON uses System.Text.Json with a source-generated context (`SettingsJsonContext`), so it needs no reflection.
- A missing, unreadable or damaged file gives the defaults; the app always starts.
- Saving writes a temporary file in the same folder and moves it over the target.
- Properties the app does not know are kept on save (`UserSettings.Other`).
- feat/app-settings adds its settings as properties of `UserSettings`.

A JSON file was chosen over `Microsoft.Windows.Storage.ApplicationData.GetForUnpackaged(publisher, product)`. That API's Microsoft Learn
page (read 2026-11-03) covers only Windows App SDK 2.0 and does not say where it stores the data.

## Smoke check
`build.ps1 -Smoke` (and so CI) checks the strings as the user sees them:
- Every key is in both `Strings/en-US` and `Strings/ru-RU`.
- The app runs twice, with `--smoke --smoke-language en` and `--smoke --smoke-language ru`. `App` passes that language to
  `LanguageService.Start` in place of the stored one, so it is applied the same way, through `PrimaryLanguageOverride` before any
  resource loads, and it is not saved.
- After the first frame, the window writes the texts it shows (`MainWindow.ShownTexts`: the title, and what `x:Uid` set on the menu and
  the empty state) to `%TEMP%\SpreadsheetEditor-smoke-texts.txt`. `build.ps1` compares each with the value in that language's
  `Resources.resw`.
- A new view adds its texts to its own `ShownTexts`.

A lookup through `ResourceManager` with a `ResourceContext` whose `Language` qualifier is `ru-RU` was tried first and dropped. A makepri
dump of the published PRI on CI (2026-11-03) showed the `ru-RU` candidates indexed, yet that lookup returned the `en-US` ones: the
context did not override the process language, which `PrimaryLanguageOverride` had set to `en-US` (why: not verified). The end-to-end
run tests what the user gets instead.

## Local check (owner, on Windows)
1. Windows display language English, no `%LOCALAPPDATA%\SpreadsheetEditor\settings.json`: start the app. The menu shows `File` › `Exit`,
   the empty state reads `No spreadsheet open`, and the title is `Spreadsheet Editor`.
2. Windows display language Russian (Settings › Time & language › Language & region › Windows display language; sign out and in): start
   the app. It shows `Файл` › `Выход` and `Таблица не открыта`, and the title is still `Spreadsheet Editor`.
3. A display language the app does not have, e.g. German: the app is in English.
4. The override, with the app closed: write `{ "language": "ru" }` to `settings.json` and start the app. It is in Russian whatever the
   Windows language is. Then try `"en"`, then `"system"`. A damaged file (e.g. `{`) starts the app in Windows' language without an error.

## Localisation: Windows run
Checked 2026-11-03 on the owner's PC, head `a7ae9bf`, built with `build.ps1` (not `-Smoke`). Windows 11 Pro 26H2, build 26300.9711;
Windows display language English (en-US); primary display 2560×1440 at 150 %; dark theme. Texts were read through UI Automation and on
screen; the File menu was opened in each run.

| Case | Result |
| --- | --- |
| 1. No `settings.json`, Windows in English | Pass: `File` › `Exit`, `No spreadsheet open` / `Open an .xlsx or .csv file to edit its rows as cards.`, title bar and taskbar thumbnail `Spreadsheet Editor`. The app did not create the file. |
| 2. `{ "language": "ru" }` | Pass: `Файл` › `Выход`, `Таблица не открыта` / `Откройте файл .xlsx или .csv, чтобы редактировать его строки как карточки.`, title still `Spreadsheet Editor`. |
| 3. `{ "language": "en" }` | Pass: English, as in case 1. |
| 4. `{ "language": "system" }` | Pass: English, as in case 1. |
| 5. `{` (damaged) | Pass: English, one window, no dialog, no .NET Runtime or Application Error event; the file was left as it was. |
| 6. Windows display language Russian | Not run (optional; it changes the Windows language). |
| 7. Look and feel | Pass: menu, menu item and empty-state texts are not cut off or overlapping in either language (150 %, dark). |

Seen in passing, outside these cases: the menu items report no access key through UI Automation, and the taskbar button's accessible name
is `SpreadsheetEditor - 1 running window`, from the exe's FileDescription (the assembly name), not `Spreadsheet Editor`.
