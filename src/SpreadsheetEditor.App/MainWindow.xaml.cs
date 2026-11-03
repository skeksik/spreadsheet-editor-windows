using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;

namespace SpreadsheetEditor.App;

public sealed partial class MainWindow : Window
{
    private bool _rendered;

    public MainWindow()
    {
        InitializeComponent();
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
