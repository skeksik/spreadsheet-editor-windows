using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>The sheet's <see cref="Sheet.ImageLayout"/> changed: width, height or pictures per row.</summary>
public sealed record PictureLayoutChanged(Sheet Sheet) : WorkbookChange(Sheet);
