namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The <c>.xlsx</c> package a workbook was read from, as the file's bytes. They are never changed: the writer opens a fresh copy for each
/// save, so two saves in a row give the same file (docs/spec.md §9.21) and whatever the app does not edit is written back as read.
/// </summary>
internal sealed class XlsxSource
{
    public XlsxSource(ReadOnlyMemory<byte> package) => Package = package;

    public ReadOnlyMemory<byte> Package { get; }
}
