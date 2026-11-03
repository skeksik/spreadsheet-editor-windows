using System.Globalization;

namespace SpreadsheetEditor.Core.Localisation;

/// <summary>
/// Which interface language the app uses (docs/spec.md §7.2 as replaced by §10.9-10.10): the user's override when it names one of the
/// app's languages, otherwise Windows' display language when the app has it, otherwise English.
/// </summary>
public static class UiLanguage
{
    /// <summary>The language tag of the English interface, as in the app's <c>Strings/en-US</c> folder.</summary>
    public const string English = "en-US";

    /// <summary>The language tag of the Russian interface, as in the app's <c>Strings/ru-RU</c> folder.</summary>
    public const string Russian = "ru-RU";

    /// <summary>The stored value of <see cref="LanguagePreference.System"/>.</summary>
    public const string SystemValue = "system";

    /// <summary>The stored value of <see cref="LanguagePreference.English"/>.</summary>
    public const string EnglishValue = "en";

    /// <summary>The stored value of <see cref="LanguagePreference.Russian"/>.</summary>
    public const string RussianValue = "ru";

    /// <summary>
    /// The language tag to use: <see cref="English"/> or <see cref="Russian"/>. With <see cref="LanguagePreference.System"/> it follows
    /// <paramref name="windowsLanguage"/> (any English or Russian culture, neutral or specific), and anything else, the invariant
    /// culture included, gives English.
    /// </summary>
    public static string Resolve(LanguagePreference preference, CultureInfo windowsLanguage)
    {
        ArgumentNullException.ThrowIfNull(windowsLanguage);
        return preference switch
        {
            LanguagePreference.English => English,
            LanguagePreference.Russian => Russian,
            _ => windowsLanguage.TwoLetterISOLanguageName switch
            {
                "ru" => Russian,
                _ => English,
            },
        };
    }

    /// <summary>
    /// The preference a stored value names, ignoring case and surrounding spaces. A missing or unknown value (a damaged or newer settings
    /// file) is <see cref="LanguagePreference.System"/>.
    /// </summary>
    public static LanguagePreference ParsePreference(string? stored) => stored?.Trim().ToLowerInvariant() switch
    {
        EnglishValue => LanguagePreference.English,
        RussianValue => LanguagePreference.Russian,
        _ => LanguagePreference.System,
    };

    /// <summary>The value stored for a preference; <see cref="ParsePreference"/> reads it back.</summary>
    public static string ToStoredValue(LanguagePreference preference) => preference switch
    {
        LanguagePreference.English => EnglishValue,
        LanguagePreference.Russian => RussianValue,
        LanguagePreference.System => SystemValue,
        _ => throw new ArgumentOutOfRangeException(nameof(preference), preference, null),
    };
}
