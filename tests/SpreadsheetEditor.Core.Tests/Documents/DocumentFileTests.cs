using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Documents;

namespace SpreadsheetEditor.Core.Tests.Documents;

[TestClass]
public sealed class DocumentFileTests
{
    private static readonly string FilePath = Path.Combine(Path.GetTempPath(), "garden", "plants.xlsx");

    [TestMethod]
    public void AnUntitledDocumentUsesTheGivenName()
    {
        var file = DocumentFile.Untitled;

        Assert.IsTrue(file.IsUntitled);
        Assert.IsNull(file.Folder);
        Assert.AreEqual("Untitled", file.DisplayName("Untitled"));
        Assert.AreEqual("Untitled", file.SuggestedFileName("Untitled"));
        Assert.AreEqual(FileFormat.Xlsx, file.Format);
        Assert.AreEqual(CsvFormat.Default, file.CsvFormatForSave);
    }

    [TestMethod]
    public void AFileShowsItsNameAndSuggestsItWithoutTheExtension()
    {
        var file = new DocumentFile(FilePath, FileFormat.Xlsx, null);

        Assert.IsFalse(file.IsUntitled);
        Assert.AreEqual(Path.GetDirectoryName(FilePath), file.Folder);
        Assert.AreEqual("plants.xlsx", file.DisplayName("Untitled"));
        Assert.AreEqual("plants", file.SuggestedFileName("Untitled"));
    }

    [TestMethod]
    public void SavedAsKeepsTheCsvFormat()
    {
        var format = new CsvFormat(CsvSeparator.Tab, CsvEncoding.Utf16LittleEndian);
        var file = new DocumentFile(null, FileFormat.Csv, format);

        var saved = file.SavedAs(FilePath, FileFormat.Xlsx);

        Assert.AreEqual(new DocumentFile(FilePath, FileFormat.Xlsx, format), saved);
        Assert.AreEqual(format, saved.CsvFormatForSave);
    }
}
