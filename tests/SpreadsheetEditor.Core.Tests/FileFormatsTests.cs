namespace SpreadsheetEditor.Core.Tests;

[TestClass]
public sealed class FileFormatsTests
{
    [TestMethod]
    [DataRow("data.csv", FileFormat.Xlsx, FileFormat.Csv)]
    [DataRow("DATA.CSV", FileFormat.Xlsx, FileFormat.Csv)]
    [DataRow("data.Csv", FileFormat.Xlsx, FileFormat.Csv)]
    [DataRow("book.xlsx", FileFormat.Csv, FileFormat.Xlsx)]
    [DataRow("BOOK.XLSX", FileFormat.Csv, FileFormat.Xlsx)]
    [DataRow("book.csv.xlsx", FileFormat.Csv, FileFormat.Xlsx)]
    [DataRow("book", FileFormat.Csv, FileFormat.Csv)]
    [DataRow("book", FileFormat.Xlsx, FileFormat.Xlsx)]
    [DataRow("book.", FileFormat.Csv, FileFormat.Csv)]
    [DataRow("notes.txt", FileFormat.Csv, FileFormat.Csv)]
    [DataRow("notes.txt", FileFormat.Xlsx, FileFormat.Xlsx)]
    [DataRow("folder.csv/book", FileFormat.Xlsx, FileFormat.Xlsx)]
    public void DecidesByExtensionIgnoringCaseOrByTheChosenFilter(string path, FileFormat chosenFilter, FileFormat expected) =>
        Assert.AreEqual(expected, FileFormats.FromPath(path, chosenFilter));
}
