using System.Text.Json;
using System.Text.Json.Serialization;

namespace SpreadsheetEditor.App.Services;

/// <summary>The settings stored per user in <see cref="SettingsStore.DefaultPath"/> (docs/notes/localisation.md, "Settings file").</summary>
internal sealed class UserSettings
{
    /// <summary>The interface language override: <c>system</c>, <c>en</c> or <c>ru</c> (Core.Localisation.UiLanguage). Null means system.</summary>
    [JsonPropertyName("language")]
    public string? Language { get; set; }

    /// <summary>Properties this version does not know (written by a newer one), kept so that saving does not drop them.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Other { get; set; }
}
