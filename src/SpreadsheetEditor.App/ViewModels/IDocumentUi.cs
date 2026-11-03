using SpreadsheetEditor.Core;

namespace SpreadsheetEditor.App.ViewModels;

/// <summary>The platform side of the File menu: the system file pickers and the dialogs. Services.DocumentUi implements it for the window.</summary>
internal interface IDocumentUi
{
    /// <summary>Shows the open picker for <c>.xlsx</c> and <c>.csv</c> files; the chosen path, or null when cancelled.</summary>
    /// <param name="folder">The folder to suggest, or null for the picker's own choice.</param>
    Task<string?> PickFileToOpenAsync(string? folder);

    /// <summary>Shows the save picker; the chosen file, or null when cancelled.</summary>
    /// <param name="suggestedName">The file name to suggest, without an extension.</param>
    /// <param name="folder">The folder to suggest, or null for the picker's own choice.</param>
    /// <param name="format">The document's format, offered first.</param>
    Task<SavePick?> PickFileToSaveAsync(string suggestedName, string? folder, FileFormat format);

    /// <summary>Asks "Do you want to save changes to {file}?" with Save, Don't save and Cancel (docs/spec.md §10.4).</summary>
    Task<UnsavedChoice> AskToSaveChangesAsync(string displayName);

    /// <summary>Shows an error (docs/spec.md §10.17): a plain message, and the technical details in an expander.</summary>
    Task ShowErrorAsync(string title, string message, string details);
}
