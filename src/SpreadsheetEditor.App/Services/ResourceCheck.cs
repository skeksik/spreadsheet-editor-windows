using Microsoft.Windows.ApplicationModel.Resources;
using SpreadsheetEditor.Core.Localisation;

namespace SpreadsheetEditor.App.Services;

/// <summary>
/// The smoke run's check of the strings (build.ps1 -Smoke, CI): every string in the app's resource file resolves in English and Russian,
/// and the Russian differs from the English unless the string is the same in every language. A Russian string that is missing falls
/// back to the English one, which is how this finds it.
/// </summary>
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
            return problems;
        }
        ResourceContext english = Context(manager, UiLanguage.English);
        ResourceContext russian = Context(manager, UiLanguage.Russian);
        for (uint index = 0; index < strings.ResourceCount; index++)
        {
            KeyValuePair<string, ResourceCandidate> en = strings.GetValueByIndex(index, english);
            KeyValuePair<string, ResourceCandidate> ru = strings.GetValueByIndex(index, russian);
            string? enText = en.Value?.ValueAsString;
            string? ruText = ru.Value?.ValueAsString;
            if (string.IsNullOrEmpty(enText))
            {
                problems.Add($"{en.Key}: no English text");
            }
            if (string.IsNullOrEmpty(ruText))
            {
                problems.Add($"{en.Key}: no Russian text");
            }
            else if (string.Equals(ruText, enText, StringComparison.Ordinal) && !Untranslated.Any(key => en.Key.EndsWith(key, StringComparison.Ordinal)))
            {
                problems.Add($"{en.Key}: the Russian text is the English one (missing from Strings/ru-RU?)");
            }
        }
        return problems;
    }

    private static ResourceContext Context(ResourceManager manager, string language)
    {
        ResourceContext context = manager.CreateResourceContext();
        context.QualifierValues["Language"] = language;
        return context;
    }
}
