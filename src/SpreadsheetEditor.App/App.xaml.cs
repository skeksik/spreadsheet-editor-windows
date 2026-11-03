using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;

namespace SpreadsheetEditor.App;

public partial class App : Application
{
    /// <summary>Started with this argument the app opens its window, waits for it to render and exits with 0 (build.ps1, CI).</summary>
    internal const string SmokeArgument = "--smoke";

    /// <summary>Where a smoke run writes an unhandled exception; build.ps1 prints it when the run fails.</summary>
    internal static string SmokeLogPath => Path.Combine(Path.GetTempPath(), "SpreadsheetEditor-smoke.log");

    private MainWindow? _window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        var smoke = Environment.GetCommandLineArgs().Contains(SmokeArgument);
        if (smoke)
        {
            UnhandledException += (_, e) => File.AppendAllText(SmokeLogPath, $"{e.Message}{Environment.NewLine}{e.Exception}{Environment.NewLine}");
        }
        _window = new MainWindow();
        if (smoke)
        {
            // Close after the first frame has finished, not inside its Rendering callback; closing the last window ends the app.
            _window.FirstRendered += (_, _) => _window.DispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, _window.Close);
        }
        _window.Activate();
    }
}
