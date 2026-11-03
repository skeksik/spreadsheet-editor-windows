namespace SpreadsheetEditor.App.ViewModels;

/// <summary>The answer to "Do you want to save changes to {file}?" (docs/spec.md §10.4).</summary>
internal enum UnsavedChoice
{
    /// <summary>Cancel, or the dialog was closed with Esc: the action that asked does not happen.</summary>
    Cancel,

    /// <summary>Save, then go on with the action if the save succeeds.</summary>
    Save,

    /// <summary>Go on with the action and lose the changes.</summary>
    DontSave,
}
