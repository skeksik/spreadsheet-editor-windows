using System.Globalization;
using System.Text;
using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Documents;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Xlsx;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Documents;

[TestClass]
public sealed class DocumentIoTests
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    private DirectoryInfo folder = null!;

    [TestInitialize]
    public void CreateFolder() => folder = Directory.CreateTempSubdirectory("document-tests-");

    [TestCleanup]
    public void DeleteFolder() => folder.Delete(recursive: true);

    [TestMethod]
    [DataRow("a.xlsx", true)]
    [DataRow("A.XLSX", true)]
    [DataRow("a.csv", true)]
    [DataRow("a.Csv", true)]
    [DataRow("a.xls", false)]
    [DataRow("a.txt", false)]
    [DataRow("a", false)]
    [DataRow("a.xlsx.png", false)]
    public void OpensOnlyXlsxAndCsvIgnoringCase(string path, bool expected) => Assert.AreEqual(expected, DocumentIo.CanOpen(path));

    [TestMethod]
    public async Task OpensACsvFileKeepingItsFormat()
    {
        var path = InFolder("plants.CSV");
        File.WriteAllBytes(path, [.. Encoding.UTF8.GetPreamble(), .. Encoding.UTF8.GetBytes("Name;Colour\r\nFern;green\r\n")]);

        var opened = await DocumentIo.OpenAsync(path, Invariant);

        Assert.AreEqual(path, opened.File.Path);
        Assert.AreEqual(FileFormat.Csv, opened.File.Format);
        Assert.AreEqual(new CsvFormat(CsvSeparator.Semicolon, CsvEncoding.Utf8WithBom), opened.File.CsvFormat);
        Assert.AreEqual("Fern", opened.Workbook.Sheets[0].Cards[0].Entries[0].Text);
    }

    [TestMethod]
    public async Task OpensAnXlsxFile()
    {
        var path = InFolder("plants.xlsx");
        File.WriteAllBytes(path, Plants().ToBytes());

        var opened = await DocumentIo.OpenAsync(path, Invariant);

        Assert.AreEqual(new DocumentFile(path, FileFormat.Xlsx, null), opened.File);
        Assert.AreEqual("Fern", opened.Workbook.Sheets[0].Cards[0].Entries[0].Text);
    }

    [TestMethod]
    public async Task AMissingCsvFileIsAReadError()
    {
        var error = await Assert.ThrowsExactlyAsync<SpreadsheetReadException>(() => DocumentIo.OpenAsync(InFolder("missing.csv"), Invariant));

        Assert.AreEqual(SpreadsheetReadError.NotFound, error.Reason);
        Assert.AreEqual(FileProblem.NotFound, FileErrors.Classify(error));
    }

    [TestMethod]
    public async Task AnEmptyCsvFileIsACsvError()
    {
        var path = InFolder("empty.csv");
        File.WriteAllBytes(path, []);

        var error = await Assert.ThrowsExactlyAsync<CsvReadException>(() => DocumentIo.OpenAsync(path, Invariant));

        Assert.AreEqual(FileProblem.CsvEmpty, FileErrors.Classify(error));
    }

    [TestMethod]
    public async Task ADamagedXlsxFileIsAReadError()
    {
        var path = InFolder("damaged.xlsx");
        File.WriteAllText(path, "not a workbook");

        var error = await Assert.ThrowsExactlyAsync<SpreadsheetReadException>(() => DocumentIo.OpenAsync(path, Invariant));

        Assert.AreEqual(FileProblem.NotAWorkbook, FileErrors.Classify(error));
    }

    [TestMethod]
    public async Task SavesXlsxAndReadsItBack()
    {
        var opened = await OpenPlants();
        opened.Workbook.Sheets[0].Cards[0].Entries[0].Text = "Moss";
        var target = InFolder("copy.xlsx");

        var result = await DocumentIo.SaveAsync(opened.Workbook, opened.Workbook.Sheets[0], opened.File, target, FileFormat.Xlsx, Invariant);

        Assert.IsTrue(result.AllSheetsWritten);
        Assert.AreEqual(new DocumentFile(target, FileFormat.Xlsx, null), result.File);
        Assert.AreEqual("Moss", XlsxReader.Read(target, Invariant).Sheets[0].Cards[0].Entries[0].Text);
    }

    [TestMethod]
    public async Task SavesTheGivenSheetAsCsvAndReportsTheOthersUnwritten()
    {
        var opened = await OpenPlants();
        var target = InFolder("tools.csv");

        var result = await DocumentIo.SaveAsync(opened.Workbook, opened.Workbook.Sheets[1], opened.File, target, FileFormat.Csv, Invariant);

        Assert.IsFalse(result.AllSheetsWritten);
        Assert.AreEqual(new DocumentFile(target, FileFormat.Csv, null), result.File);
        Assert.AreEqual("Tool\r\nSpade\r\n", File.ReadAllText(target));
    }

    [TestMethod]
    public async Task SavesASingleSheetAsCsvInTheFormatItWasReadWith()
    {
        var source = InFolder("plants.csv");
        var bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("Name\tColour\r\nFern\tgreen\r\n")).ToArray();
        File.WriteAllBytes(source, bytes);
        var opened = await DocumentIo.OpenAsync(source, Invariant);
        var target = InFolder("again.csv");

        var result = await DocumentIo.SaveAsync(opened.Workbook, null, opened.File, target, FileFormat.Csv, Invariant);

        Assert.IsTrue(result.AllSheetsWritten);
        Assert.AreEqual(opened.File.CsvFormat, result.File.CsvFormat);
        CollectionAssert.AreEqual(bytes, File.ReadAllBytes(target));
    }

    [TestMethod]
    public async Task ACsvDocumentSavedAsXlsxKeepsItsCsvFormatForALaterCsvSave()
    {
        var source = InFolder("plants.csv");
        File.WriteAllText(source, "Name;Colour\r\nFern;green\r\n");
        var opened = await DocumentIo.OpenAsync(source, Invariant);

        var result = await DocumentIo.SaveAsync(opened.Workbook, null, opened.File, InFolder("plants.xlsx"), FileFormat.Xlsx, Invariant);

        Assert.AreEqual(FileFormat.Xlsx, result.File.Format);
        Assert.AreEqual(CsvSeparator.Semicolon, result.File.CsvFormatForSave.Separator);
    }

    [TestMethod]
    public async Task AFailedSaveLeavesTheFileUnchanged()
    {
        var opened = await OpenPlants();
        var target = InFolder("plants.csv");
        File.WriteAllText(target, "old\r\n");
        opened.Workbook.Sheets[0].Cards[0].Entries[0].Text = "\uD800";

        var error = await Assert.ThrowsExactlyAsync<SpreadsheetWriteException>(
            () => DocumentIo.SaveAsync(opened.Workbook, null, opened.File, target, FileFormat.Csv, Invariant));

        Assert.AreEqual(FileProblem.WriteFailed, FileErrors.Classify(error));
        Assert.AreEqual("old\r\n", File.ReadAllText(target));
    }

    [TestMethod]
    public async Task SavingAsCsvASheetOfAnotherWorkbookIsRefused()
    {
        var opened = await OpenPlants();

        await Assert.ThrowsExactlyAsync<ArgumentException>(
            () => DocumentIo.SaveAsync(opened.Workbook, new Sheet("Other"), opened.File, InFolder("x.csv"), FileFormat.Csv, Invariant));
    }

    [TestMethod]
    public void NewHasOneDefaultSheetAndNoFile()
    {
        var created = DocumentIo.New("Sheet");

        Assert.AreSame(DocumentFile.Untitled, created.File);
        var sheet = created.Workbook.Sheets.Single();
        Assert.AreEqual("Sheet1", sheet.Name);
        Assert.AreEqual(1, sheet.Columns.Count);
        Assert.AreEqual(1, sheet.Cards.Count);
        Assert.AreEqual(ImageLayout.Default, sheet.ImageLayout);
    }

    [TestMethod]
    [DataRow("Sheet")]
    [DataRow("Лист")]
    [DataRow("A base name far longer than thirty-one characters")]
    public void NewNamesItsSheetUnderTheSheetNameRules(string baseName)
    {
        var name = DocumentIo.New(baseName).Workbook.Sheets[0].Name;

        Assert.AreEqual(SheetNameCheck.Valid, SheetNames.Check(name, []));
    }

    [TestMethod]
    public async Task ANewDocumentSavesAsXlsx()
    {
        var created = DocumentIo.New("Sheet");
        created.Workbook.Sheets[0].Cards[0].Entries[0].Text = "Fern";
        var target = InFolder("new.xlsx");

        await DocumentIo.SaveAsync(created.Workbook, null, created.File, target, FileFormat.Xlsx, Invariant);

        Assert.AreEqual("Sheet1", XlsxReader.Read(target, Invariant).Sheets[0].Name);
    }

    [TestMethod]
    public async Task AnUntouchedNewDocumentReopensWithItsSheet()
    {
        var created = DocumentIo.New("Sheet");
        var target = InFolder("untouched.xlsx");

        await DocumentIo.SaveAsync(created.Workbook, null, created.File, target, FileFormat.Xlsx, Invariant);
        var reopened = await DocumentIo.OpenAsync(target, Invariant);

        Assert.AreEqual("Sheet1", reopened.Workbook.Sheets.Single().Name);
    }

    private string InFolder(string name) => Path.Combine(folder.FullName, name);

    private async Task<Core.Documents.OpenedDocument> OpenPlants()
    {
        var path = InFolder("plants.xlsx");
        File.WriteAllBytes(path, Plants().ToBytes());
        return await DocumentIo.OpenAsync(path, Invariant);
    }

    private static TestWorkbook Plants()
    {
        var workbook = new TestWorkbook();
        workbook.AddSheet("Plants").Row("A1", "Name", "Colour").Row("A2", "Fern", "green");
        workbook.AddSheet("Tools").Row("A1", "Tool").Row("A2", "Spade");
        return workbook;
    }
}
