using System.Text.Json;

namespace SpreadsheetEditor.App.Services;

/// <summary>
/// The user's settings in a JSON file, <see cref="DefaultPath"/>. A missing, unreadable or damaged file gives the defaults and never stops
/// the app from starting; saving writes a temporary file next to the target and then moves it over the target, so a failed save leaves
/// the previous file whole. feat/app-settings adds its settings to <see cref="UserSettings"/>.
/// </summary>
internal sealed class SettingsStore
{
    private readonly string _path;

    private SettingsStore(string path, UserSettings current)
    {
        _path = path;
        Current = current;
    }

    /// <summary><c>%LOCALAPPDATA%\SpreadsheetEditor\settings.json</c>.</summary>
    public static string DefaultPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SpreadsheetEditor", "settings.json");

    /// <summary>The settings as last loaded or saved.</summary>
    public UserSettings Current { get; private set; }

    /// <summary>Reads the settings from <paramref name="path"/>, or the defaults when it is missing, unreadable or not valid JSON.</summary>
    public static SettingsStore Load(string path)
    {
        UserSettings? settings = null;
        try
        {
            using FileStream stream = File.OpenRead(path);
            settings = JsonSerializer.Deserialize(stream, SettingsJsonContext.Default.UserSettings);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or JsonException or NotSupportedException)
        {
            // FileNotFoundException and DirectoryNotFoundException are IOExceptions: a first start has no file yet.
        }
        return new SettingsStore(path, settings ?? new UserSettings());
    }

    /// <summary>Writes <paramref name="settings"/> and makes them <see cref="Current"/>. Throws IOException or UnauthorizedAccessException.</summary>
    public void Save(UserSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        string directory = Path.GetDirectoryName(_path)!;
        Directory.CreateDirectory(directory);
        string temporary = Path.Combine(directory, $"{Path.GetFileName(_path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (FileStream stream = File.Create(temporary))
            {
                JsonSerializer.Serialize(stream, settings, SettingsJsonContext.Default.UserSettings);
            }
            File.Move(temporary, _path, overwrite: true);
        }
        finally
        {
            File.Delete(temporary);
        }
        Current = settings;
    }
}
