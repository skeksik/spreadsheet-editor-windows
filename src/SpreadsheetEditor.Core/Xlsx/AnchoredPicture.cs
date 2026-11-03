using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>A picture anchored to a worksheet cell: the 1-based row and column of its top-left anchor cell, and the picture.</summary>
internal sealed record AnchoredPicture(int Row, int Column, EntryImage Image);
