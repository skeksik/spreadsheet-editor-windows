namespace SpreadsheetEditor.Core.Localisation;

/// <summary>The user's choice of interface language (docs/spec.md §10.9): follow Windows, or one of the app's languages.</summary>
public enum LanguagePreference
{
    /// <summary>Use Windows' display language when the app has it, otherwise English.</summary>
    System,

    /// <summary>Always English.</summary>
    English,

    /// <summary>Always Russian.</summary>
    Russian,
}
