using System.Collections.ObjectModel;
using System.Globalization;
using SpreadsheetEditor.App.Services;
using SpreadsheetEditor.Core;
using SpreadsheetEditor.Core.Documents;
using SpreadsheetEditor.Core.Editing;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.App.ViewModels;

/// <summary>
/// The window's document and its File menu (docs/spec.md §2.1-2.2 as replaced by §10.4-10.8, §10.17): New, Open, Save, Save as, opening
/// a dropped file, and the unsaved-changes question before any of them and before closing. Reading and writing run on the thread pool
/// (core DocumentIo) while <see cref="IsBusy"/> is true and the File commands are disabled (§9.53).
/// </summary>
public sealed partial class MainViewModel : ObservableObject
{
    private readonly IDocumentUi _ui;
    private readonly CultureInfo _culture;
    private readonly WorkbookEditor _editor = new(new Workbook());

    // The open document's file, or null while no document is open.
    private DocumentFile? _file;

    private bool _isBusy;

    // True while a File flow runs (a question, a picker, a read or a write): another command, a drop or a close waits for none and is
    // refused, so two flows never overlap (and two ContentDialogs are never open at once).
    private bool _inFlow;

    // The last folder a file was opened from or saved to, for this session only (persisting it is feat/app-settings).
    private string? _lastFolder;

    internal MainViewModel(IDocumentUi ui, CultureInfo culture)
    {
        _ui = ui;
        _culture = culture;
        _editor.StateChanged += (_, _) => RefreshTitle();
        Title = Strings.Get("AppDisplayName");
    }

    /// <summary>The window title: "Spreadsheet Editor", or "{file} - Spreadsheet Editor" with "*" before the name while modified (§10.8).</summary>
    public string Title { get; private set; }

    /// <summary>True while a document is open.</summary>
    public bool HasDocument => _file is not null;

    /// <summary>True while no document is open: the empty state shows.</summary>
    public bool HasNoDocument => _file is null;

    /// <summary>True while a file is read or written: a progress bar shows and the File commands are disabled (§9.53).</summary>
    public bool IsBusy => _isBusy;

    /// <summary>New, Open and Exit are enabled unless a file is being read or written.</summary>
    public bool CanRunFileCommands => !_isBusy;

    /// <summary>Save and Save as are enabled with a document open, unless a file is being read or written.</summary>
    public bool CanSave => HasDocument && !_isBusy;

    /// <summary>
    /// The placeholder for the open document until the sheet tabs and cards exist (feat/app-sheets, feat/app-cards): one line per sheet
    /// with its card count.
    /// </summary>
    public ObservableCollection<string> SheetSummaries { get; } = [];

    /// <summary>True when a dropped file may be opened now: no other File flow runs.</summary>
    internal bool CanAcceptDrop => !_inFlow && !_isBusy;

    /// <summary>True when closing must first ask about unsaved changes, or wait for a running flow (the close is then refused).</summary>
    internal bool NeedsCloseConfirmation => _inFlow || _isBusy || NeedsSaveQuestion;

    private bool NeedsSaveQuestion => HasDocument && _editor.IsModified;

    private static string UntitledName => Strings.Get("UntitledName");

    // The sheet a .csv save writes (§9.34). Until feat/app-sheets selects sheets, the current sheet is the first.
    private Sheet? CurrentSheet => _editor.Workbook.Sheets.FirstOrDefault();

    /// <summary>File › New (Ctrl+N): asks about unsaved changes, then opens a new untitled document with one empty sheet.</summary>
    public Task NewAsync() => RunFlowAsync(async () =>
    {
        if (!await ResolveUnsavedChangesAsync())
        {
            return;
        }
        Show(DocumentIo.New(Strings.Get("NewSheetBaseName")));
    });

    /// <summary>File › Open (Ctrl+O): asks about unsaved changes, then shows the open picker and reads the chosen file.</summary>
    public Task OpenAsync() => RunFlowAsync(async () =>
    {
        if (!await ResolveUnsavedChangesAsync())
        {
            return;
        }
        var path = await _ui.PickFileToOpenAsync(_lastFolder ?? _file?.Folder);
        if (path is not null)
        {
            await OpenFileAsync(path);
        }
    });

    /// <summary>A <c>.xlsx</c> or <c>.csv</c> file dropped on the window: asks about unsaved changes, then reads it.</summary>
    internal Task OpenDroppedAsync(string path) => RunFlowAsync(async () =>
    {
        if (!await ResolveUnsavedChangesAsync())
        {
            return;
        }
        await OpenFileAsync(path);
    });

    /// <summary>File › Save (Ctrl+S): writes the current file; a document that was never saved asks for a path (§10.5).</summary>
    public Task SaveAsync() => RunFlowAsync(async () => await SaveCoreAsync());

    /// <summary>File › Save as (Ctrl+Shift+S): asks for a path and format, then writes there.</summary>
    public Task SaveAsAsync() => RunFlowAsync(async () => await SaveAsCoreAsync());

    /// <summary>
    /// Before the window closes (the close button, Alt+F4, File › Exit): true when it may close. While a file is read or written, or another
    /// flow runs, it is false and nothing is asked; otherwise unsaved changes are asked about first.
    /// </summary>
    internal async Task<bool> ConfirmCloseAsync()
    {
        if (_inFlow || _isBusy)
        {
            return false;
        }
        _inFlow = true;
        try
        {
            return await ResolveUnsavedChangesAsync();
        }
        finally
        {
            _inFlow = false;
        }
    }

    private async Task RunFlowAsync(Func<Task> flow)
    {
        if (_inFlow || _isBusy)
        {
            return;
        }
        _inFlow = true;
        try
        {
            await flow();
        }
        finally
        {
            _inFlow = false;
        }
    }

    /// <summary>
    /// True when the action that asked may go on: there were no unsaved changes, the user chose Don't save, or chose Save and the save
    /// succeeded. A .csv save that left other sheets unwritten keeps the document modified (§9.34) and still counts as saved here, so the
    /// action goes on instead of asking again.
    /// </summary>
    private async Task<bool> ResolveUnsavedChangesAsync()
    {
        if (!NeedsSaveQuestion)
        {
            return true;
        }
        _editor.EndTyping();
        return await _ui.AskToSaveChangesAsync(_file!.DisplayName(UntitledName)) switch
        {
            UnsavedChoice.Save => await SaveCoreAsync(),
            UnsavedChoice.DontSave => true,
            _ => false,
        };
    }

    private Task<bool> SaveCoreAsync()
    {
        if (_file is null)
        {
            return Task.FromResult(false);
        }
        return _file.Path is null ? SaveAsCoreAsync() : SaveToAsync(_file.Path, _file.Format);
    }

    private async Task<bool> SaveAsCoreAsync()
    {
        if (_file is null)
        {
            return false;
        }
        var pick = await _ui.PickFileToSaveAsync(_file.SuggestedFileName(UntitledName), _file.Folder ?? _lastFolder, _file.Format);
        if (pick is null)
        {
            return false;
        }
        return await SaveToAsync(pick.Path, FileFormats.FromPath(pick.Path, pick.FallbackFormat));
    }

    private async Task<bool> SaveToAsync(string path, FileFormat format)
    {
        _editor.EndTyping();
        SaveResult result;
        SetBusy(true);
        try
        {
            result = await DocumentIo.SaveAsync(_editor.Workbook, CurrentSheet, _file!, path, format, _culture);
        }
        catch (Exception ex)
        {
            // Every failure is shown, never thrown: the document stays open and modified (§10.17).
            SetBusy(false);
            await ShowErrorAsync("SaveFailedTitle", path, ex);
            return false;
        }
        SetBusy(false);
        _lastFolder = Path.GetDirectoryName(path);
        _file = result.File;
        if (result.AllSheetsWritten)
        {
            _editor.MarkSaved();
        }
        RefreshTitle();
        return true;
    }

    private async Task OpenFileAsync(string path)
    {
        OpenedDocument opened;
        SetBusy(true);
        try
        {
            opened = await DocumentIo.OpenAsync(path, _culture);
        }
        catch (Exception ex)
        {
            // The open document stays as it was (§9.13); a bad file never crashes the app (§10.17).
            SetBusy(false);
            await ShowErrorAsync("OpenFailedTitle", path, ex);
            return;
        }
        SetBusy(false);
        _lastFolder = opened.File.Folder;
        Show(opened);
    }

    private Task ShowErrorAsync(string titleKey, string path, Exception error) =>
        _ui.ShowErrorAsync(Strings.Format(titleKey, Path.GetFileName(path)), FileProblemText.For(FileErrors.Classify(error)), error.Message);

    private void Show(OpenedDocument document)
    {
        var hadDocument = HasDocument;
        _file = document.File;
        _editor.Load(document.Workbook);
        SheetSummaries.Clear();
        foreach (var sheet in document.Workbook.Sheets)
        {
            SheetSummaries.Add(Strings.Format("SheetSummary", sheet.Name, sheet.Cards.Count));
        }
        if (document.Workbook.Sheets.Count == 0)
        {
            SheetSummaries.Add(Strings.Get("NoSheetsSummary"));
        }
        if (!hadDocument)
        {
            OnPropertyChanged(nameof(HasDocument));
            OnPropertyChanged(nameof(HasNoDocument));
            OnPropertyChanged(nameof(CanSave));
        }
        RefreshTitle();
    }

    private void SetBusy(bool busy)
    {
        _isBusy = busy;
        OnPropertyChanged(nameof(IsBusy));
        OnPropertyChanged(nameof(CanRunFileCommands));
        OnPropertyChanged(nameof(CanSave));
    }

    private void RefreshTitle()
    {
        var title = WindowTitles.Compose(
            _file?.DisplayName(UntitledName), _editor.IsModified, Strings.Get("WindowTitleWithFile"), Strings.Get("AppDisplayName"));
        if (title != Title)
        {
            Title = title;
            OnPropertyChanged(nameof(Title));
        }
    }
}
