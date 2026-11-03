using System.IO.Compression;
using System.Text;
using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Model;
using static SpreadsheetEditor.Core.Tests.Csv.CsvTestFiles;

namespace SpreadsheetEditor.Core.Tests.Csv;

[TestClass]
public sealed class CsvReaderTests
{
    [TestMethod]
    [DataRow("a,b,c\r\n1,2,3\r\n", CsvSeparator.Comma, "a|b|c¶1|2|3")]
    [DataRow("a;b;c\r\n1;2;3\r\n", CsvSeparator.Semicolon, "a|b|c¶1|2|3")]
    [DataRow("a\tb\tc\r\n1\t2\t3\r\n", CsvSeparator.Tab, "a|b|c¶1|2|3")]
    [DataRow("a;b,c;d\r\n", CsvSeparator.Semicolon, "a|b,c|d¶||")]
    [DataRow("\"x,y\";z\r\n", CsvSeparator.Semicolon, "x,y|z¶|")]
    [DataRow("a,b;c\r\n", CsvSeparator.Comma, "a|b;c¶|")]
    [DataRow("a;b\tc\r\n", CsvSeparator.Semicolon, "a|b\tc¶|")]
    [DataRow("a\tb,c\r\n", CsvSeparator.Comma, "a\tb|c¶|")]
    [DataRow("single\r\nx;y\r\n", CsvSeparator.Comma, "single¶x;y")]
    [DataRow("\"a\r\nb,c\";d\r\n", CsvSeparator.Semicolon, "a\r\nb,c|d¶|")]
    [DataRow("\r\n\r\na;b\r\n1,5;2,5\r\n", CsvSeparator.Semicolon, "a|b¶1,5|2,5")]
    [DataRow("a\"b,c\";d\r\n", CsvSeparator.Comma, "a\"b|c\";d¶|")]
    public void DetectsTheSeparatorFromTheHeader(string text, CsvSeparator expected, string sheet)
    {
        var result = Read(text);

        Assert.AreEqual(expected, result.Format.Separator);
        Assert.AreEqual(sheet, Describe(result.Workbook.Sheets[0]));
    }

    [TestMethod]
    [DataRow(CsvEncoding.Utf8)]
    [DataRow(CsvEncoding.Utf8WithBom)]
    [DataRow(CsvEncoding.Utf16LittleEndian)]
    [DataRow(CsvEncoding.Utf16BigEndian)]
    public void DetectsTheEncodingFromTheByteOrderMark(CsvEncoding encoding)
    {
        var result = Read(Encode("Name;Город\r\nÄpfel;Ёж 🍎\r\n", encoding));

        Assert.AreEqual(encoding, result.Format.Encoding);
        Assert.AreEqual(CsvSeparator.Semicolon, result.Format.Separator);
        Assert.AreEqual("Name|Город¶Äpfel|Ёж 🍎", Describe(result.Workbook.Sheets[0]));
    }

    [TestMethod]
    public void RejectsALegacyCodePage()
    {
        // "Name\r\nCafé\r\n" in Windows-1252: the é is a lone 0xE9, which is not UTF-8.
        byte[] bytes = [.. Encoding.ASCII.GetBytes("Name\r\nCaf"), 0xE9, .. Encoding.ASCII.GetBytes("\r\n")];

        var error = Assert.ThrowsExactly<CsvReadException>(() => Read(bytes));

        Assert.AreEqual(CsvReadFailure.UnsupportedEncoding, error.Reason);
    }

    [TestMethod]
    public void RejectsUtf16WithoutAByteOrderMark()
    {
        var bytes = new UnicodeEncoding(false, false).GetBytes("Name,Value\r\na,b\r\n");

        var error = Assert.ThrowsExactly<CsvReadException>(() => Read(bytes));

        Assert.AreEqual(CsvReadFailure.UnsupportedEncoding, error.Reason);
    }

    [TestMethod]
    public void RejectsInvalidUtf16()
    {
        // A lone high surrogate (D800) after the byte order mark.
        byte[] bytes = [0xFF, 0xFE, (byte)'a', 0, (byte)'\r', 0, (byte)'\n', 0, 0x00, 0xD8, (byte)'\r', 0, (byte)'\n', 0];

        var error = Assert.ThrowsExactly<CsvReadException>(() => Read(bytes));

        Assert.AreEqual(CsvReadFailure.UnsupportedEncoding, error.Reason);
    }

    [TestMethod]
    [DataRow(new byte[] { 0xFF, 0xFE, 0x00, 0x00, (byte)'a', 0, 0, 0 })]
    [DataRow(new byte[] { 0x00, 0x00, 0xFE, 0xFF, 0, 0, 0, (byte)'a' })]
    public void RejectsUtf32(byte[] bytes)
    {
        var error = Assert.ThrowsExactly<CsvReadException>(() => Read(bytes));

        Assert.AreEqual(CsvReadFailure.UnsupportedEncoding, error.Reason);
    }

    [TestMethod]
    public void ReadsQuotedValues()
    {
        var sheet = ReadSheet("\"a \"\"quoted\"\" word\",\"x,y\",\"\"\r\n\"\",plain \"inner\" quote,\"closed\"after\r\n");

        Assert.AreEqual("a \"quoted\" word|x,y|¶|plain \"inner\" quote|closedafter", Describe(sheet));
    }

    [TestMethod]
    public void KeepsLineBreaksInsideQuotedValuesAsRead()
    {
        var sheet = ReadSheet("Name,Note\r\nA,\"line 1\r\nline 2\nline 3\rline 4\"\r\nB,x\r\n");

        Assert.AreEqual(2, sheet.Cards.Count);
        Assert.AreEqual("line 1\r\nline 2\nline 3\rline 4", sheet.Cards[0].Entries[1].Text);
        Assert.AreEqual("B|x", string.Join('|', sheet.Cards[1].Entries.Select(e => e.Text)));
    }

    [TestMethod]
    public void EndsRecordsAtCrLfLfOrCr()
    {
        var sheet = ReadSheet("a,b\r1,2\n3,4\r\n5,6");

        Assert.AreEqual("a|b¶1|2¶3|4¶5|6", Describe(sheet));
    }

    [TestMethod]
    public void SkipsBlankLines()
    {
        var sheet = ReadSheet("\r\na,b\r\n\r\n1,2\r\n\n\r3,4\r\n\r\n");

        Assert.AreEqual("a|b¶1|2¶3|4", Describe(sheet));
    }

    [TestMethod]
    public void ReadsLinesOfEmptyValuesAsCards()
    {
        Assert.AreEqual("a|b¶|¶1|2", Describe(ReadSheet("a,b\r\n,\r\n1,2\r\n")));
        Assert.AreEqual("a¶¶x", Describe(ReadSheet("a\r\n\"\"\r\nx\r\n")));
    }

    [TestMethod]
    public void PadsRaggedRecords()
    {
        var sheet = ReadSheet("a,b\r\n1\r\n1,2,3\r\n");

        Assert.AreEqual("a|b|¶1||¶1|2|3", Describe(sheet));
    }

    [TestMethod]
    [DataRow("", CsvEncoding.Utf8)]
    [DataRow("", CsvEncoding.Utf8WithBom)]
    [DataRow("", CsvEncoding.Utf16LittleEndian)]
    [DataRow("", CsvEncoding.Utf16BigEndian)]
    [DataRow("\r\n\n\r", CsvEncoding.Utf8)]
    public void RejectsAFileWithoutAHeader(string text, CsvEncoding encoding)
    {
        var error = Assert.ThrowsExactly<CsvReadException>(() => Read(Encode(text, encoding)));

        Assert.AreEqual(CsvReadFailure.Empty, error.Reason);
        Assert.IsFalse(string.IsNullOrEmpty(error.Message));
    }

    [TestMethod]
    public void RejectsAnUnterminatedQuote()
    {
        var error = Assert.ThrowsExactly<CsvReadException>(() => Read("a,b\r\n1,\"open\r\n2,3\r\n"));

        Assert.AreEqual(CsvReadFailure.UnterminatedQuote, error.Reason);
        Assert.AreEqual(2, error.LineNumber);
    }

    [TestMethod]
    public void RejectsAnUnterminatedQuoteInTheHeader()
    {
        var error = Assert.ThrowsExactly<CsvReadException>(() => Read("a,\"b\r\n1,2\r\n"));

        Assert.AreEqual(CsvReadFailure.UnterminatedQuote, error.Reason);
        Assert.AreEqual(1, error.LineNumber);
    }

    [TestMethod]
    [DataRow("a,b\r\n")]
    [DataRow("a,b")]
    public void GivesAHeaderOnlyFileOneEmptyCard(string text)
    {
        var sheet = ReadSheet(text);

        Assert.AreEqual(1, sheet.Cards.Count);
        Assert.AreEqual(2, sheet.Cards[0].Entries.Count);
        foreach (var entry in sheet.Cards[0].Entries)
        {
            Assert.AreEqual("", entry.Text);
            Assert.IsNull(entry.Original);
        }
    }

    [TestMethod]
    public void KeepsValuesAsText()
    {
        var sheet = ReadSheet("Date,Number,Flag,Spaces\r\n2024-01-15 meeting notes,0012,TRUE,  padded  \r\n2024-01-15,1e3,false,\r\n");

        Assert.AreEqual("Date|Number|Flag|Spaces¶2024-01-15 meeting notes|0012|TRUE|  padded  ¶2024-01-15|1e3|false|", Describe(sheet));
        var first = sheet.Cards[0].Entries[0];
        Assert.AreEqual(new CellValue(CellValueKind.Text, "2024-01-15 meeting notes"), first.Original);
        Assert.AreEqual(new CellValue(CellValueKind.Text, "1e3"), sheet.Cards[1].Entries[1].Original);
        Assert.IsNull(sheet.Cards[1].Entries[3].Original);
    }

    [TestMethod]
    public void KeepsColumnNamesAsRead()
    {
        var sheet = ReadSheet(",a,a,Unnamed: 3,\r\n1,2,3,4,5\r\n");

        Assert.AreEqual("|a|a|Unnamed: 3|", string.Join('|', sheet.Columns.Select(c => c.Name)));
    }

    [TestMethod]
    public void BuildsOneSheetWithTheDefaultLayout()
    {
        var result = Read("a,b\r\n1,2\r\n3,4\r\n");

        Assert.AreEqual(1, result.Workbook.Sheets.Count);
        Assert.IsNull(result.Workbook.SourcePath);
        var sheet = result.Workbook.Sheets[0];
        Assert.AreEqual("data.csv", sheet.Name);
        Assert.IsNull(sheet.TableName);
        Assert.AreEqual(1, sheet.HeaderRow);
        Assert.AreEqual(2, sheet.FirstDataRow);
        Assert.AreEqual(1, sheet.FirstColumn);
        Assert.AreEqual(ImageLayout.Default, sheet.ImageLayout);
        Assert.IsFalse(sheet.Columns.Any(c => c.IsImageColumn));
        Assert.IsTrue(sheet.Cards.All(c => c.SourceRowSpan == 1));
        Assert.IsTrue(sheet.Cards.SelectMany(c => c.Entries).All(e => e.Images.Count == 0 && e.Style is null));
    }

    [TestMethod]
    public void ReadsAFileNamedWithItsExtension()
    {
        var folder = Directory.CreateTempSubdirectory("csv-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "people.CSV");
            File.WriteAllBytes(path, Encode("Name\r\nAda\r\n", CsvEncoding.Utf8WithBom));

            var result = CsvReader.Read(path);

            Assert.AreEqual("people.CSV", result.Workbook.Sheets[0].Name);
            Assert.AreEqual(path, result.Workbook.SourcePath);
            Assert.AreEqual(new CsvFormat(CsvSeparator.Comma, CsvEncoding.Utf8WithBom), result.Format);
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }

    [TestMethod]
    public void ReadsAStreamThatCannotSeek()
    {
        using var compressed = new MemoryStream();
        using (var gzip = new GZipStream(compressed, CompressionMode.Compress, leaveOpen: true))
        {
            gzip.Write(Encode("a;b\r\n1;2\r\n", CsvEncoding.Utf16BigEndian));
        }
        compressed.Position = 0;
        using var stream = new GZipStream(compressed, CompressionMode.Decompress);
        Assert.IsFalse(stream.CanSeek);

        var result = CsvReader.Read(stream, "data.csv");

        Assert.AreEqual(new CsvFormat(CsvSeparator.Semicolon, CsvEncoding.Utf16BigEndian), result.Format);
        Assert.AreEqual("a|b¶1|2", Describe(result.Workbook.Sheets[0]));
    }

    [TestMethod]
    public void ReadsValuesLongerThanTheBuffer()
    {
        var longValue = new string('x', 100_000) + "\r\n" + new string('y', 50_000);
        var sheet = ReadSheet("a,b\r\n\"" + longValue + "\",z\r\n");

        Assert.AreEqual(longValue, sheet.Cards[0].Entries[0].Text);
        Assert.AreEqual("z", sheet.Cards[0].Entries[1].Text);
    }

    [TestMethod]
    [DataRow("a\r\n1\r\n", true)]
    [DataRow("a\n1\n", true)]
    [DataRow("a\r1\r", true)]
    [DataRow("a\r\n1\r\n\r\n", true)]
    [DataRow("a\r\n1", false)]
    [DataRow("a", false)]
    [DataRow("a\r\n\"1\r\n2\"", false)]
    public void RecordsWhetherTheFileEndsWithALineBreak(string text, bool expected) =>
        Assert.AreEqual(expected, Read(text).Format.FinalLineBreak);
}
