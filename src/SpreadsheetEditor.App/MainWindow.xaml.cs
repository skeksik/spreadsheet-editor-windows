using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
using SpreadsheetEditor.App.Services;

namespace SpreadsheetEditor.App;

public sealed partial class MainWindow : Window
{
    private bool _rendered;

    public MainWindow()
    {
        InitializeComponent();
        // Window is not a DependencyObject, so x:Uid cannot set its title (not verified); the title bar shows the same text. With a
        // document open the title becomes WindowTitleWithFile, "{file} - Spreadsheet Editor" (spec §10.8, feat/app-document).
        Title = Strings.Get("AppDisplayName");
        AppTitleBar.Title = Title;
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        CompositionTarget.Rendering += OnRendering;
    }

    /// <summary>Raised once, after the window's first frame (App.SmokeArgument).</summary>
    public event EventHandler? FirstRendered;

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

    private void OnExitClick(object sender, RoutedEventArgs e) => Close();
}
