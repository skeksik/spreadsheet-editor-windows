using System.Globalization;
using Microsoft.Windows.ApplicationModel.Resources;

namespace SpreadsheetEditor.App.Services;

/// <summary>
/// The app's strings in code, from <c>Strings/&lt;lang&gt;/Resources.resw</c> through MRT Core (docs/notes/localisation.md). XAML sets
/// its strings with x:Uid instead. Use keys without a dot here; a key with one belongs to an x:Uid.
/// </summary>
internal static class Strings
{
    // The "Resources" subtree of the app's main resource map, in the language LanguageService.Start chose.
    private static readonly Lazy<ResourceLoader> Loader = new(() => new ResourceLoader());

    /// <summary>The string <paramref name="key"/> in the interface language.</summary>
    public static string Get(string key) => Loader.Value.GetString(key);

    /// <summary>
    /// The string <paramref name="key"/> with its <c>{0}</c>-style placeholders filled. Numbers in the arguments take Windows' regional
    /// format (CurrentCulture), as CA1305 asks; cell values are formatted by the core with LanguageService.UiCulture before they get here.
    /// </summary>
    public static string Format(string key, params object?[] args) => string.Format(CultureInfo.CurrentCulture, Get(key), args);
}
