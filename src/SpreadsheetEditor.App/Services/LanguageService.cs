using System.Globalization;
using Microsoft.Windows.Globalization;
using SpreadsheetEditor.Core.Localisation;

namespace SpreadsheetEditor.App.Services;

/// <summary>
/// The interface language (docs/spec.md §10.9-10.10, docs/notes/localisation.md): Windows' display language when it is English or
/// Russian, otherwise English, unless the user chose one. The choice is stored in <see cref="SettingsStore"/> and takes effect at the
/// next start.
/// </summary>
internal sealed class LanguageService
{
    private readonly SettingsStore _settings;

    private LanguageService(SettingsStore settings, LanguagePreference preference, CultureInfo uiCulture)
    {
        _settings = settings;
        Preference = preference;
        UiCulture = uiCulture;
    }

    /// <summary>The choice this run started with; <see cref="SavedPreference"/> is the one the next start uses.</summary>
    public LanguagePreference Preference { get; }

    /// <summary>The choice stored for the next start.</summary>
    public LanguagePreference SavedPreference => UiLanguage.ParsePreference(_settings.Current.Language);

    /// <summary>True when the stored choice gives another language than this run's, so a restart is needed to see it.</summary>
    public bool RestartNeeded => UiLanguage.Resolve(SavedPreference, WindowsLanguage) != UiCulture.Name;

    /// <summary>The interface language of this run (en-US or ru-RU). Later tasks pass it to the core's value formatting (Values.CellText).</summary>
    public CultureInfo UiCulture { get; }

    /// <summary>Windows' display language, as the process started with it.</summary>
    public static CultureInfo WindowsLanguage { get; } = CultureInfo.CurrentUICulture;

    /// <summary>
    /// Chooses this run's language and applies it. Call once, at startup, before any resource is loaded (before InitializeComponent).
    /// </summary>
    /// <remarks>
    /// Microsoft Learn, "ApplicationLanguages.PrimaryLanguageOverride Property (Microsoft.Windows.Globalization)", Windows App SDK 1.6-2.0
    /// (read 2026-11-03): for an unpackaged app the override "is not persisted between app sessions and should be set each time the app
    /// is loaded"; "Set PrimaryLanguageOverride setting during app load, before any resource is loaded"; the example says resources
    /// "loaded manually with ResourceLoader and XAML resources fetched with x:Uid" follow it. It is set every start, also when following
    /// Windows, so that a display language the app does not have gives English.
    /// <see cref="WindowsLanguage"/> is the culture the process starts with: Microsoft Learn, "CultureInfo.CurrentUICulture Property"
    /// (.NET 10, read 2026-11-03): "On systems that use the Windows operating system, the common language runtime calls the Windows
    /// GetUserDefaultUILanguage function to set the current UI culture. GetUserDefaultUILanguage returns the default UI culture set by
    /// the user."
    /// </remarks>
    public static LanguageService Start(SettingsStore settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        LanguagePreference preference = UiLanguage.ParsePreference(settings.Current.Language);
        string tag = UiLanguage.Resolve(preference, WindowsLanguage);
        ApplicationLanguages.PrimaryLanguageOverride = tag;
        CultureInfo culture = CultureInfo.GetCultureInfo(tag);
        // The interface language only; CurrentCulture (Windows' regional format) is left as it is.
        CultureInfo.DefaultThreadCurrentUICulture = culture;
        CultureInfo.CurrentUICulture = culture;
        return new LanguageService(settings, preference, culture);
    }

    /// <summary>
    /// Stores the choice for the next start; this run keeps its language (see <see cref="Start"/>: the override is read when resources
    /// load). Throws IOException or UnauthorizedAccessException when the settings file cannot be written.
    /// </summary>
    public void SavePreference(LanguagePreference preference)
    {
        UserSettings current = _settings.Current;
        _settings.Save(new UserSettings { Language = UiLanguage.ToStoredValue(preference), Other = current.Other });
    }
}
