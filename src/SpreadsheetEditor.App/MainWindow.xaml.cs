using System.ComponentModel;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Media;
using SpreadsheetEditor.App.Services;
using SpreadsheetEditor.App.ViewModels;
using SpreadsheetEditor.Core.Documents;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;

namespace SpreadsheetEditor.App;

public sealed partial class MainWindow : Window
{
    private bool _rendered;

    // Set once the unsaved-changes question allowed the close, so the close that follows is not asked about again.
    private bool _closeConfirmed;

    public MainWindow()
    {
        ViewModel = new MainViewModel(new DocumentUi(this), App.InterfaceLanguage.UiCulture);
        InitializeComponent();
        // Window is not a DependencyObject, so x:Uid cannot set its title (not verified); the title bar shows the same text: "Spreadsheet
        // Editor", or "{file} - Spreadsheet Editor" with "*" before the name while there are unsaved changes (spec §10.8).
        ApplyTitle();
        ViewModel.PropertyChanged += OnViewModelPropertyChanged;
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        AppWindow.Closing += OnClosing;
        CompositionTarget.Rendering += OnRendering;
    }

    /// <summary>Raised once, after the window's first frame (App.SmokeArgument).</summary>
    public event EventHandler? FirstRendered;

    /// <summary>The document and the File menu.</summary>
    public MainViewModel ViewModel { get; }

    private void OnRendering(object? sender, object e)
    {
        if (_rendered)
        {
            return;
        }
        _rendered = true;
        CompositionTarget.Rendering -= OnRendering;
        FirstRendered?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// The texts the window shows, by their resource key (App.SmokeTextsPath): what x:Uid and Services.Strings set, so the smoke run checks
    /// the language the user sees. Access keys are included: they are localised too.
    /// </summary>
    internal IEnumerable<KeyValuePair<string, string>> ShownTexts()
    {
        yield return new("AppDisplayName", Title);
        yield return new("MainWindow_FileMenu.Title", FileMenu.Title);
        yield return new("MainWindow_FileMenu.AccessKey", FileMenu.AccessKey);
        yield return new("MainWindow_NewItem.Text", NewItem.Text);
        yield return new("MainWindow_NewItem.AccessKey", NewItem.AccessKey);
        yield return new("MainWindow_OpenItem.Text", OpenItem.Text);
        yield return new("MainWindow_OpenItem.AccessKey", OpenItem.AccessKey);
        yield return new("MainWindow_SaveItem.Text", SaveItem.Text);
        yield return new("MainWindow_SaveItem.AccessKey", SaveItem.AccessKey);
        yield return new("MainWindow_SaveAsItem.Text", SaveAsItem.Text);
        yield return new("MainWindow_SaveAsItem.AccessKey", SaveAsItem.AccessKey);
        yield return new("MainWindow_ExitItem.Text", ExitItem.Text);
        yield return new("MainWindow_ExitItem.AccessKey", ExitItem.AccessKey);
        yield return new("MainWindow_BusyBar.[using:Microsoft.UI.Xaml.Automation]AutomationProperties.Name", AutomationProperties.GetName(BusyBar));
        yield return new("MainWindow_EmptyTitle.Text", EmptyTitle.Text);
        yield return new("MainWindow_EmptyHint.Text", EmptyHint.Text);
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainViewModel.Title))
        {
            ApplyTitle();
        }
    }

    private void ApplyTitle()
    {
        Title = ViewModel.Title;
        AppTitleBar.Title = ViewModel.Title;
    }

    private async void OnNewClick(object sender, RoutedEventArgs e) => await ViewModel.NewAsync();

    private async void OnOpenClick(object sender, RoutedEventArgs e) => await ViewModel.OpenAsync();

    private async void OnSaveClick(object sender, RoutedEventArgs e) => await ViewModel.SaveAsync();

    private async void OnSaveAsClick(object sender, RoutedEventArgs e) => await ViewModel.SaveAsAsync();

    // File › Exit takes the same way as the close button.
    private async void OnExitClick(object sender, RoutedEventArgs e) => await CloseAfterConfirmingAsync();

    /// <remarks>
    /// Microsoft Learn, "AppWindow.Closing Event" (Windows App SDK 1.0-2.0, read 2026-11-03): it "Occurs when a window is being closed
    /// through a system affordance" and "does not occur when the AppWindow.Destroy method is called". The close is cancelled while the
    /// question is open and repeated once it allows the close; with nothing to ask, the window closes at once (the smoke run's close).
    /// </remarks>
    private async void OnClosing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (_closeConfirmed || !ViewModel.NeedsCloseConfirmation)
        {
            return;
        }
        args.Cancel = true;
        await CloseAfterConfirmingAsync();
    }

    private async Task CloseAfterConfirmingAsync()
    {
        // False while a file is read or written or another question is open: the close is refused and nothing else is started.
        if (await ViewModel.ConfirmCloseAsync())
        {
            _closeConfirmed = true;
            Close();
        }
    }

    /// <remarks>
    /// Reading the dragged files is asynchronous, so the handlers take a deferral. Microsoft Learn, "DragEventArgs.GetDeferral Method"
    /// (Windows App SDK 0.8-2.0, read 2026-11-03): it "Supports asynchronous drag-and-drop operations by creating and returning a
    /// DragOperationDeferral object."
    /// </remarks>
    private async void OnDragOver(object sender, DragEventArgs e)
    {
        if (!ViewModel.CanAcceptDrop || !e.DataView.Contains(StandardDataFormats.StorageItems))
        {
            e.AcceptedOperation = DataPackageOperation.None;
            return;
        }
        var deferral = e.GetDeferral();
        try
        {
            e.AcceptedOperation = await FirstOpenableFileAsync(e.DataView) is null ? DataPackageOperation.None : DataPackageOperation.Copy;
        }
        finally
        {
            deferral.Complete();
        }
    }

    private async void OnDrop(object sender, DragEventArgs e)
    {
        if (!ViewModel.CanAcceptDrop || !e.DataView.Contains(StandardDataFormats.StorageItems))
        {
            return;
        }
        string? path;
        var deferral = e.GetDeferral();
        try
        {
            path = await FirstOpenableFileAsync(e.DataView);
        }
        finally
        {
            deferral.Complete();
        }
        // The question and the read come after the drop has completed, so the source app is not kept waiting.
        if (path is not null)
        {
            await ViewModel.OpenDroppedAsync(path);
        }
    }

    // The first dragged item, when it is a .xlsx or .csv file; other files are ignored (spec §2.1: only the first file is used).
    private static async Task<string?> FirstOpenableFileAsync(DataPackageView data)
    {
        try
        {
            var items = await data.GetStorageItemsAsync();
            return items.Count > 0 && items[0].IsOfType(StorageItemTypes.File) && DocumentIo.CanOpen(items[0].Path) ? items[0].Path : null;
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or InvalidOperationException or System.Runtime.InteropServices.COMException)
        {
            // A source that cannot hand over its items is treated like one with nothing to open.
            return null;
        }
    }
}
