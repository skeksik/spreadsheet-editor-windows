using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using SpreadsheetEditor.App.Services;
using SpreadsheetEditor.Core.Localisation;

namespace SpreadsheetEditor.App;

public partial class App : Application
{
    /// <summary>Started with this argument the app opens its window, waits for it to render and exits with 0 (build.ps1, CI).</summary>
    internal const string SmokeArgument = "--smoke";

    /// <summary>
    /// With <see cref="SmokeArgument"/>, the interface language of the run (<c>en</c> or <c>ru</c>) in place of the stored one; it is
    /// applied like the stored override, before any resource loads, and not saved. build.ps1 runs the smoke once per language.
    /// </summary>
    internal const string SmokeLanguageArgument = "--smoke-language";

    /// <summary>Where a smoke run writes an unhandled exception; build.ps1 prints it when the run fails.</summary>
    internal static string SmokeLogPath => Path.Combine(Path.GetTempPath(), "SpreadsheetEditor-smoke.log");

    /// <summary>
    /// Where a smoke run writes the texts its window showed, one <c>key&lt;TAB&gt;text</c> line each (MainWindow.ShownTexts); build.ps1
    /// compares them with <c>Strings/&lt;lang&gt;/Resources.resw</c>.
    /// </summary>
    internal static string SmokeTextsPath => Path.Combine(Path.GetTempPath(), "SpreadsheetEditor-smoke-texts.txt");

    private MainWindow? _window;

    public App()
    {
        // Before InitializeComponent: the interface language must be set before any resource loads (LanguageService.Start).
        Settings = SettingsStore.Load(SettingsStore.DefaultPath);
        InterfaceLanguage = LanguageService.Start(Settings, SmokeLanguage(Environment.GetCommandLineArgs()));
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
        }
        _window = new MainWindow();
        if (smoke)
        {
            // Close after the first frame has finished, not inside its Rendering callback; closing the last window ends the app.
            _window.FirstRendered += (_, _) =>
            {
                File.WriteAllLines(SmokeTextsPath, _window.ShownTexts().Select(pair => $"{pair.Key}\t{pair.Value}"));
                _window.DispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, _window.Close);
            };
        }
        _window.Activate();
    }

    /// <summary>The language <see cref="SmokeLanguageArgument"/> names in a smoke run, otherwise null (the stored choice applies).</summary>
    private static LanguagePreference? SmokeLanguage(string[] arguments)
    {
        int index = Array.IndexOf(arguments, SmokeLanguageArgument);
        if (!arguments.Contains(SmokeArgument) || index < 0 || index + 1 >= arguments.Length)
        {
            return null;
        }
        return UiLanguage.ParsePreference(arguments[index + 1]);
    }
}
