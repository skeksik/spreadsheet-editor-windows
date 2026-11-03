using System.Text;
using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;
using static SpreadsheetEditor.Core.Tests.Csv.CsvTestFiles;

namespace SpreadsheetEditor.Core.Tests.Csv;

[TestClass]
public sealed class CsvWriterTests
{
    private const string Awkward = "Name|Note¶plain|has,comma¶has \"quote\"|two\nlines¶semi;colon|tab\there¶ spaced |cr\ronly";

    [TestMethod]
    public void QuotesOnlyWhereNeededWithCommas()
    {
        Assert.AreEqual(
            "Name,Note\r\nplain,\"has,comma\"\r\n\"has \"\"quote\"\"\",\"two\nlines\"\r\nsemi;colon,tab\there\r\n spaced ,\"cr\ronly\"\r\n",
            WriteText(MakeSheet(Awkward)));
    }

    [TestMethod]
    public void QuotesOnlyWhereNeededWithSemicolons()
    {
        Assert.AreEqual(
            "Name;Note\r\nplain;has,comma\r\n\"has \"\"quote\"\"\";\"two\nlines\"\r\n\"semi;colon\";tab\there\r\n spaced ;\"cr\ronly\"\r\n",
            WriteText(MakeSheet(Awkward), CsvSeparator.Semicolon));
    }

    [TestMethod]
    public void QuotesOnlyWhereNeededWithTabs()
    {
        Assert.AreEqual(
            "Name\tNote\r\nplain\thas,comma\r\n\"has \"\"quote\"\"\"\t\"two\nlines\"\r\nsemi;colon\t\"tab\there\"\r\n spaced \t\"cr\ronly\"\r\n",
            WriteText(MakeSheet(Awkward), CsvSeparator.Tab));
    }

    [TestMethod]
    public void OmitsTheHeaderLineWhenHeadersAreHidden()
    {
        var sheet = MakeSheet("0|1¶a|b¶c|d");
        sheet.HeaderRow = null;
        sheet.FirstDataRow = 1;

        Assert.AreEqual("a,b\r\nc,d\r\n", WriteText(sheet));
    }

    [TestMethod]
    public void WritesAnEmptySingleValueAsQuotes()
    {
        Assert.AreEqual("Name\r\n\"\"\r\nx\r\n", WriteText(MakeSheet("Name¶¶x")));
        Assert.AreEqual("\"\"\r\nx\r\n", WriteText(MakeSheet("¶x")));
    }

    [TestMethod]
    public void WritesTheHeaderAloneWhenTheOnlyCardIsEmpty()
    {
        Assert.AreEqual("a,b\r\n", WriteText(MakeSheet("a|b¶|")));
        Assert.AreEqual("a\r\n", WriteText(MakeSheet("a¶")));
    }

    [TestMethod]
    public void WritesEveryCardWhenThereAreSeveral()
    {
        Assert.AreEqual("a,b\r\n,\r\n,\r\n", WriteText(MakeSheet("a|b¶|¶|")));
        Assert.AreEqual("a,b\r\n,\r\n1,2\r\n", WriteText(MakeSheet("a|b¶|¶1|2")));
    }

    [TestMethod]
    public void WritesAnEmptyNameAsAnEmptyHeader()
    {
        Assert.AreEqual(",b,b\r\n1,2,3\r\n", WriteText(MakeSheet("|b|b¶1|2|3")));
    }

    [TestMethod]
    public void WritesNothingForASheetWithoutColumns()
    {
        var sheet = new Sheet("Empty");
        sheet.Cards.Add(new Card());

        Assert.AreEqual(0, Write(sheet, new CsvFormat(CsvSeparator.Comma, CsvEncoding.Utf8)).Length);
    }

    [TestMethod]
    public void WritesOnlyTheText()
    {
        var sheet = MakeSheet("Photo|Name¶caption|Ada");
        sheet.TableName = "People";
        sheet.Columns[0].IsImageColumn = true;
        sheet.Cards[0].SourceRowSpan = 3;
        sheet.Cards[0].Entries[0].Images.Add(new EntryImage(new byte[3], ImageFormat.Png, 10, 20));
        sheet.Cards[0].Entries[1].Style = new CellStyle { Fill = new RgbColor(255, 0, 0) };

        Assert.AreEqual("Photo,Name\r\ncaption,Ada\r\n", WriteText(sheet));
    }

    [TestMethod]
    public void WritesMissingEntriesAsEmptyValues()
    {
        var sheet = MakeSheet("a|b|c¶1|2|3");
        sheet.Cards.Add(new Card());
        sheet.Cards[1].Entries.Add(new Entry { Text = "4" });

        Assert.AreEqual("a,b,c\r\n1,2,3\r\n4,,\r\n", WriteText(sheet));
    }

    [TestMethod]
    [DataRow(CsvEncoding.Utf8, "")]
    [DataRow(CsvEncoding.Utf8WithBom, "EFBBBF")]
    [DataRow(CsvEncoding.Utf16LittleEndian, "FFFE")]
    [DataRow(CsvEncoding.Utf16BigEndian, "FEFF")]
    public void WritesTheEncodingWithItsByteOrderMark(CsvEncoding encoding, string byteOrderMarkHex)
    {
        var byteOrderMark = Convert.FromHexString(byteOrderMarkHex);
        var bytes = Write(MakeSheet("Name¶Ёж 🍎"), new CsvFormat(CsvSeparator.Comma, encoding));

        CollectionAssert.AreEqual(Encode("Name\r\nЁж 🍎\r\n", encoding), bytes);
        CollectionAssert.AreEqual(byteOrderMark, bytes.Take(byteOrderMark.Length).ToArray());
    }

    [TestMethod]
    public void ExportsWithCommaAndUtf8WithoutAByteOrderMarkByDefault()
    {
        Assert.AreEqual(new CsvFormat(CsvSeparator.Comma, CsvEncoding.Utf8), CsvFormat.Default);
        var bytes = Write(MakeSheet("a|b¶é|2"), CsvFormat.Default);

        CollectionAssert.AreEqual(Encoding.UTF8.GetBytes("a,b\r\né,2\r\n"), bytes);
    }

    [TestMethod]
    public void LeavesTheStreamOpen()
    {
        using var stream = new MemoryStream();

        CsvWriter.Write(MakeSheet("a¶1"), stream, CsvFormat.Default);

        Assert.IsTrue(stream.CanWrite);
        Assert.AreEqual(6, stream.Length);
    }

    [TestMethod]
    public void WritesAFile()
    {
        var folder = Directory.CreateTempSubdirectory("csv-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "out.csv");
            File.WriteAllText(path, "old content that is longer than the new one");

            CsvWriter.Write(MakeSheet("a|b¶1|2"), path, new CsvFormat(CsvSeparator.Semicolon, CsvEncoding.Utf8WithBom));

            CollectionAssert.AreEqual(Encode("a;b\r\n1;2\r\n", CsvEncoding.Utf8WithBom), File.ReadAllBytes(path));
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }

    [TestMethod]
    public void WritingAFileLeavesNoTemporaryFile()
    {
        var folder = Directory.CreateTempSubdirectory("csv-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "out.csv");

            CsvWriter.Write(MakeSheet("a¶1"), path, CsvFormat.Default);

            Assert.AreEqual("out.csv", folder.GetFiles().Single().Name);
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }

    [TestMethod]
    public void AFailedWriteLeavesTheTargetUnchanged()
    {
        var folder = Directory.CreateTempSubdirectory("csv-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "out.csv");
            File.WriteAllText(path, "a\r\nold\r\n");
            var original = File.ReadAllBytes(path);
            var sheet = MakeSheet("a¶x");
            sheet.Cards[0].Entries[0].Text = "lone \uD800 surrogate";

            var error = Assert.ThrowsExactly<SpreadsheetWriteException>(() => CsvWriter.Write(sheet, path, CsvFormat.Default));

            Assert.AreEqual(SpreadsheetWriteError.WriteFailed, error.Reason);
            CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
            Assert.AreEqual(1, folder.GetFiles().Length);
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }

    // Built in code: an attribute stores its strings as UTF-8, which turns a lone surrogate into U+FFFD.
    [TestMethod]
    [DataRow(0)]
    [DataRow(1)]
    [DataRow(2)]
    public void ALoneSurrogateIsAWriteErrorAndWritesNothing(int index)
    {
        var text = new[] { "lone \uD800 high", "lone \uDC00 low", "ends high \uD83D" }[index];
        var sheet = MakeSheet("a¶x");
        sheet.Cards[0].Entries[0].Text = text;
        using var stream = new MemoryStream();

        var error = Assert.ThrowsExactly<SpreadsheetWriteException>(() => CsvWriter.Write(sheet, stream, CsvFormat.Default));

        Assert.AreEqual(SpreadsheetWriteError.WriteFailed, error.Reason);
        Assert.AreEqual(0, stream.Length);
    }

    [TestMethod]
    public void WritesAPairedSurrogate()
    {
        Assert.AreEqual("a\r\n\U0001F33F\r\n", WriteText(MakeSheet("a¶\U0001F33F")));
    }

    [TestMethod]
    public void AMissingFolderIsAWriteError()
    {
        var folder = Directory.CreateTempSubdirectory("csv-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "missing", "out.csv");

            var error = Assert.ThrowsExactly<SpreadsheetWriteException>(() => CsvWriter.Write(MakeSheet("a¶1"), path, CsvFormat.Default));

            Assert.AreEqual(SpreadsheetWriteError.WriteFailed, error.Reason);
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }

    [TestMethod]
    public async Task WriteAsyncWritesTheFile()
    {
        var folder = Directory.CreateTempSubdirectory("csv-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "out.csv");

            await CsvWriter.WriteAsync(MakeSheet("a¶1"), path, CsvFormat.Default);

            Assert.AreEqual("a\r\n1\r\n", File.ReadAllText(path));
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }
}
