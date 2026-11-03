using Microsoft.Windows.ApplicationModel.Resources;
using Microsoft.Windows.Globalization;
using SpreadsheetEditor.Core.Localisation;

namespace SpreadsheetEditor.App.Services;

/// <summary>
/// The smoke run's check of the strings (build.ps1 -Smoke, CI): every string in the app's resource file resolves in English and Russian,
/// and the Russian differs from the English unless the string is the same in every language. A Russian string that is missing falls
/// back to the English one, which is how this finds it.
/// </summary>
/// <remarks>
/// The lookup follows Microsoft Learn, "Localize strings in your UI and app package manifest", section "Load a string for a specific
/// language or other context" (read 2026-11-03): <c>CreateResourceContext</c>, <c>QualifierValues["Language"]</c>,
/// <c>MainResourceMap.GetSubtree("Resources")</c>, <c>GetValue(name, context)</c>.
/// </remarks>
internal static class ResourceCheck
{
    /// <summary>Keys whose text is the same in every language: the product name (spec §7.2) and language names (spec §10.9).</summary>
    private static readonly string[] Untranslated = ["AppDisplayName", "WindowTitleWithFile", "LanguageName_English", "LanguageName_Russian"];

    /// <summary>What is wrong, one line per problem; empty when every string resolves in both languages.</summary>
    public static IReadOnlyList<string> FindProblems()
    {
        var problems = new List<string>();
        var manager = new ResourceManager();
        ResourceMap? strings = manager.MainResourceMap.TryGetSubtree("Resources");
        if (strings is null || strings.ResourceCount == 0)
        {
            problems.Add("the app's resource map has no strings (Strings/<lang>/Resources.resw)");
            problems.AddRange(Diagnostics(manager, strings, null));
            return problems;
        }
        ResourceContext english = Context(manager, UiLanguage.English);
        ResourceContext russian = Context(manager, UiLanguage.Russian);
        string? firstFailing = null;
        for (uint index = 0; index < strings.ResourceCount; index++)
        {
            string name = strings.GetValueByIndex(index).Key;
            string? enText = strings.TryGetValue(name, english)?.ValueAsString;
            string? ruText = strings.TryGetValue(name, russian)?.ValueAsString;
            if (string.IsNullOrEmpty(enText))
            {
                problems.Add($"{name}: no English text");
            }
            if (string.IsNullOrEmpty(ruText))
            {
                problems.Add($"{name}: no Russian text");
            }
            else if (string.Equals(ruText, enText, StringComparison.Ordinal) && !IsUntranslated(name))
            {
                problems.Add($"{name}: the Russian text is the English one (missing from Strings/ru-RU?)");
                firstFailing ??= name;
            }
        }
        if (problems.Count > 0)
        {
            problems.AddRange(Diagnostics(manager, strings, firstFailing));
        }
        return problems;
    }

    private static bool IsUntranslated(string name) => Untranslated.Any(key => name.EndsWith(key, StringComparison.Ordinal));

    private static ResourceContext Context(ResourceManager manager, string language)
    {
        ResourceContext context = manager.CreateResourceContext();
        context.QualifierValues["Language"] = language;
        return context;
    }

    /// <summary>
    /// Temporary (feat/app-localisation, until the smoke check passes on CI): what the resource system sees, to tell a PRI without the
    /// Russian strings from a lookup that ignores the language. Written to the smoke log only when the check fails.
    /// </summary>
    private static IEnumerable<string> Diagnostics(ResourceManager manager, ResourceMap? strings, string? name)
    {
        yield return "diagnostics:";
        yield return $"  PrimaryLanguageOverride = '{ApplicationLanguages.PrimaryLanguageOverride}'";
        yield return $"  ApplicationLanguages.Languages = [{string.Join(", ", ApplicationLanguages.Languages)}]";
        yield return $"  ApplicationLanguages.ManifestLanguages = [{string.Join(", ", ApplicationLanguages.ManifestLanguages)}]";
        foreach (string pri in Directory.EnumerateFiles(AppContext.BaseDirectory, "*.pri"))
        {
            byte[] bytes = File.ReadAllBytes(pri);
            yield return $"  {Path.GetFileName(pri)}: {bytes.Length} bytes, Cyrillic text: UTF-16 {HasCyrillicUtf16(bytes)}, UTF-8 {HasCyrillicUtf8(bytes)}";
        }
        ResourceContext russian = Context(manager, UiLanguage.Russian);
        yield return $"  ru context QualifierValues = {Dump(russian.QualifierValues)}";
        yield return $"  default context QualifierValues = {Dump(manager.CreateResourceContext().QualifierValues)}";
        if (strings is null || name is null)
        {
            yield break;
        }
        yield return $"  '{name}' via QualifierValues ru-RU: '{strings.TryGetValue(name, russian)?.ValueAsString}'";
        ResourceContext neutral = manager.CreateResourceContext();
        neutral.QualifierValues["Language"] = "ru";
        yield return $"  '{name}' via QualifierValues ru: '{strings.TryGetValue(name, neutral)?.ValueAsString}'";
        yield return $"  '{name}' via ResourceLoader (process language): '{new ResourceLoader().GetString(name)}'";
    }

    private static string Dump(IDictionary<string, string> values) =>
        string.Join("; ", values.Select(pair => $"{pair.Key}={pair.Value}"));

    // Two or more Cyrillic letters (U+0410-U+044F) in a row, little-endian UTF-16.
    private static bool HasCyrillicUtf16(byte[] bytes)
    {
        for (int i = 0; i + 3 < bytes.Length; i++)
        {
            if (bytes[i + 1] == 0x04 && bytes[i] is >= 0x10 and <= 0x4F && bytes[i + 3] == 0x04 && bytes[i + 2] is >= 0x10 and <= 0x4F)
            {
                return true;
            }
        }
        return false;
    }

    // Two or more Cyrillic letters (U+0410-U+044F) in a row, UTF-8 (lead byte D0 or D1).
    private static bool HasCyrillicUtf8(byte[] bytes)
    {
        for (int i = 0; i + 3 < bytes.Length; i++)
        {
            if (bytes[i] is 0xD0 or 0xD1 && bytes[i + 1] is >= 0x80 and <= 0xBF && bytes[i + 2] is 0xD0 or 0xD1 && bytes[i + 3] is >= 0x80 and <= 0xBF)
            {
                return true;
            }
        }
        return false;
    }
}
