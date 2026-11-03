using System.Globalization;

namespace SpreadsheetEditor.Core.Documents;

/// <summary>The window title (docs/spec.md §10.8).</summary>
public static class WindowTitles
{
    /// <summary>
    /// The marker before the file name while the document has unsaved changes, as Notepad shows it: Windows Insider Blog, "Announcing
    /// Windows 10 Insider Preview Build 18298" (2018-12-10, read 2026-11-03): "Notepad will now show an asterisk (*) in the title bar
    /// before the title of a document" "to indicate that a document has unsaved changes." Not verified: that today's Windows 11 Notepad
    /// still shows it this way.
    /// </summary>
    public const string UnsavedMarker = "*";

    /// <summary>
    /// The title: <paramref name="appName"/> with no document; otherwise <paramref name="titleWithFile"/> with the file name in place of
    /// <c>{0}</c>, marked with <see cref="UnsavedMarker"/> when <paramref name="modified"/>.
    /// </summary>
    /// <param name="displayName">The document's name (<see cref="DocumentFile.DisplayName"/>), or null when no document is open.</param>
    /// <param name="modified">Whether the document has unsaved changes.</param>
    /// <param name="titleWithFile">The localised "{0} - Spreadsheet Editor".</param>
    /// <param name="appName">The localised app name.</param>
    public static string Compose(string? displayName, bool modified, string titleWithFile, string appName)
    {
        ArgumentNullException.ThrowIfNull(titleWithFile);
        ArgumentNullException.ThrowIfNull(appName);
        if (displayName is null)
        {
            return appName;
        }
        return string.Format(CultureInfo.InvariantCulture, titleWithFile, modified ? UnsavedMarker + displayName : displayName);
    }
}
