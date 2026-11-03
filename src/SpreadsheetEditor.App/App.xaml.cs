using Microsoft.UI.Xaml;

namespace SpreadsheetEditor.App;

public partial class App : Application
{
    /// <summary>Started with this argument the app opens its window, waits for it to render and exits with 0 (build.ps1, CI).</summary>
    internal const string SmokeArgument = "--smoke";

    private MainWindow? _window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        var smoke = Environment.GetCommandLineArgs().Contains(SmokeArgument);
        _window = new MainWindow();
        if (smoke)
        {
            _window.FirstRendered += (_, _) => Exit();
        }
        _window.Activate();
    }
}
