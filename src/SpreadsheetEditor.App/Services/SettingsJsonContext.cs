using System.Text.Json.Serialization;

namespace SpreadsheetEditor.App.Services;

/// <summary>Source-generated JSON for <see cref="UserSettings"/>: no reflection, so the app stays ready for trimming and NativeAOT.</summary>
[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(UserSettings))]
internal sealed partial class SettingsJsonContext : JsonSerializerContext;
