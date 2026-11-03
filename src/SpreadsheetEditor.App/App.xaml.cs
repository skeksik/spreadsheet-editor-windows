using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using SpreadsheetEditor.App.Services;

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
        // Before InitializeComponent: the interface language must be set before any resource loads (LanguageService.Start).
        Settings = SettingsStore.Load(SettingsStore.DefaultPath);
        InterfaceLanguage = LanguageService.Start(Settings);
        InitializeComponent();
    }

    /// <summary>The user's stored settings.</summary>
    internal static SettingsStore Settings { get; private set; } = null!;

    /// <summary>The interface language of this run, and the stored choice for the next start.</summary>
    internal static LanguageService InterfaceLanguage { get; private set; } = null!;

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        var smoke = Environment.GetCommandLineArgs().Contains(SmokeArgument);
        if (smoke)
        {
            UnhandledException += (_, e) => File.AppendAllText(SmokeLogPath, $"{e.Message}{Environment.NewLine}{e.Exception}{Environment.NewLine}");
            IReadOnlyList<string> problems = ResourceCheck.FindProblems();
            if (problems.Count > 0)
            {
                File.AppendAllLines(SmokeLogPath, problems.Prepend("strings:"));
                Environment.Exit(1);
            }
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
