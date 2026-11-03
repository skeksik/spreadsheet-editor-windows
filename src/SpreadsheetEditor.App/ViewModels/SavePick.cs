using SpreadsheetEditor.Core;

namespace SpreadsheetEditor.App.ViewModels;

/// <summary>The file the user chose in the save picker.</summary>
/// <param name="Path">The chosen path.</param>
/// <param name="FallbackFormat">
/// The format to use when the path has no <c>.xlsx</c> or <c>.csv</c> extension (docs/spec.md §9.35): the document's format, not the file
/// type the user chose, which the picker does not report (see Services.DocumentUi).
/// </param>
internal sealed record SavePick(string Path, FileFormat FallbackFormat);
