using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.Windows.Storage.Pickers;
using SpreadsheetEditor.App.ViewModels;
using SpreadsheetEditor.Core;

namespace SpreadsheetEditor.App.Services;

/// <summary>
/// The system file pickers and the File menu's dialogs for one window (docs/spec.md §10.3-10.5, §10.17).
/// </summary>
/// <remarks>
/// The pickers are the Windows App SDK's, which an unpackaged app uses with its window's id. Microsoft Learn, read 2026-11-03,
/// "FileOpenPicker Class" and "FileSavePicker Class" (Microsoft.Windows.Storage.Pickers, Windows App SDK 1.8 and 2.0): the constructor takes
/// a WindowId "to reference the parent for the file picker dialog"; "FileSavePicker.PickSaveFileAsync Method": it "Returns a PickFileResult
/// object that contains the path of the saved file" and "Returns null if the file dialog was cancelled or closed without selecting a saved
/// file"; "FileSavePicker.SuggestedFolder Property": "This property is only a suggestion ... If not specified or the specified path doesn't
/// exist, defaults to the last folder the user visited." The classes are in Microsoft.WindowsAppSDK.Foundation 2.3.12
/// (metadata/Microsoft.Windows.Storage.Pickers.winmd), already referenced. Seen in the owner's check (docs/notes/app-document-check.md,
/// 2026-11-03): the save picker lists FileTypeChoices in the order they were added, and appends the chosen type's extension to a name
/// typed without one.
/// </remarks>
internal sealed class DocumentUi(Window window) : IDocumentUi
{
    private const string XlsxExtension = ".xlsx";
    private const string CsvExtension = ".csv";

    public async Task<string?> PickFileToOpenAsync(string? folder)
    {
        // Only the formats the app reads; .xls is not offered (docs/spec.md §9.11, manager's choice 2026-11-03). Named choices, as the
        // save picker has: Microsoft Learn, "FileOpenPicker.FileTypeChoices Property" (Windows App SDK 2.0, read 2026-11-03): "A dictionary
        // that groups file types into labeled choices that the file open picker dialog presents to the user"; "FileTypeChoices takes
        // precedence over FileTypeFilter when both defined." The picker adds the patterns to each label, as the save picker showed
        // "Excel workbook (*.xlsx)" in the owner's check (docs/notes/app-document-check.md); not verified for the open picker.
        var picker = new FileOpenPicker(window.AppWindow.Id);
        picker.FileTypeChoices.Add(Strings.Get("OpenTypeSpreadsheets"), new List<string> { XlsxExtension, CsvExtension });
        picker.FileTypeChoices.Add(Strings.Get("SaveTypeXlsx"), new List<string> { XlsxExtension });
        picker.FileTypeChoices.Add(Strings.Get("SaveTypeCsv"), new List<string> { CsvExtension });
        if (folder is not null)
        {
            picker.SuggestedFolder = folder;
        }
        var result = await picker.PickSingleFileAsync();
        return result?.Path;
    }

    public async Task<SavePick?> PickFileToSaveAsync(string suggestedName, string? folder, FileFormat format)
    {
        var picker = new FileSavePicker(window.AppWindow.Id) { SuggestedFileName = suggestedName };
        if (folder is not null)
        {
            picker.SuggestedFolder = folder;
        }
        FileFormat[] order = format == FileFormat.Csv ? [FileFormat.Csv, FileFormat.Xlsx] : [FileFormat.Xlsx, FileFormat.Csv];
        foreach (var choice in order)
        {
            picker.FileTypeChoices.Add(
                Strings.Get(choice == FileFormat.Csv ? "SaveTypeCsv" : "SaveTypeXlsx"),
                new List<string> { choice == FileFormat.Csv ? CsvExtension : XlsxExtension });
        }
        var result = await picker.PickSaveFileAsync();
        // The result holds only the path: Microsoft Learn, "PickFileResult Class" (read 2026-11-03) lists Path as its only property ("a lightweight class that contains a string attribute representing the file path"), and
        // the picker has no property for the chosen file type after the pick. So the document's format is the fallback for a path without
        // an extension. The picker appends the chosen type's extension to a name typed without one (the owner's check), so the fallback
        // is rare.
        return result is null ? null : new SavePick(result.Path, format);
    }

    public async Task<UnsavedChoice> AskToSaveChangesAsync(string displayName)
    {
        var dialog = NewDialog();
        dialog.Title = Strings.Format("SaveChangesQuestion", displayName);
        dialog.PrimaryButtonText = Strings.Get("SaveChangesSave");
        dialog.SecondaryButtonText = Strings.Get("SaveChangesDontSave");
        dialog.CloseButtonText = Strings.Get("SaveChangesCancel");
        dialog.DefaultButton = ContentDialogButton.Primary;
        return await dialog.ShowAsync() switch
        {
            ContentDialogResult.Primary => UnsavedChoice.Save,
            ContentDialogResult.Secondary => UnsavedChoice.DontSave,
            _ => UnsavedChoice.Cancel,
        };
    }

    public async Task ShowErrorAsync(string title, string message, string details)
    {
        var panel = new StackPanel { Spacing = 12 };
        panel.Children.Add(new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(new Expander
        {
            Header = Strings.Get("ErrorDetailsHeader"),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Left,
            Content = new TextBlock { Text = details, TextWrapping = TextWrapping.Wrap, IsTextSelectionEnabled = true },
        });
        var dialog = NewDialog();
        dialog.Title = title;
        dialog.Content = panel;
        dialog.CloseButtonText = Strings.Get("ErrorDialogClose");
        dialog.DefaultButton = ContentDialogButton.Close;
        // Enter chooses OK even though focus starts on the Details expander, whose header would take Enter (Space still expands it).
        // The dialog gives initial focus to the first focusable element in its content and only then to the default button, and its
        // buttons are not tab stops: microsoft/microsoft-ui-xaml (main, read 2026-11-03), dxaml/xcp/dxaml/lib/ContentDialog_Partial.cpp,
        // SetInitialFocusElement ("Try to set focus to the first focusable element in the content area", then "try to focus the default
        // button"), and controls/dev/CommonStyles/ContentDialog_themeresources.xaml (CloseButton IsTabStop="False"). Not verified: that
        // PreviewKeyDown on the dialog sees Enter before the expander does.
        dialog.PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Windows.System.VirtualKey.Enter && !e.KeyStatus.WasKeyDown)
            {
                e.Handled = true;
                dialog.Hide();
            }
        };
        await dialog.ShowAsync();
    }

    // A ContentDialog needs the XamlRoot of the window it shows over; DefaultContentDialogStyle gives it the Fluent look, as the WinUI 3
    // Gallery's ContentDialog sample sets it (not verified against a Microsoft Learn page).
    private ContentDialog NewDialog() => new()
    {
        XamlRoot = window.Content.XamlRoot,
        Style = Application.Current.Resources["DefaultContentDialogStyle"] as Style,
    };
}
