using SpreadsheetEditor.Core.Csv;
using static SpreadsheetEditor.Core.Tests.Csv.CsvTestFiles;

namespace SpreadsheetEditor.Core.Tests.Csv;

[TestClass]
public sealed class CsvRoundTripTests
{
    /// <summary>Files as the writer writes them: CRLF after every line, quotes only where needed.</summary>
    private static readonly string[] CanonicalFiles =
    [
        "Name,Note\r\nAda,plain\r\nBo,\"has,comma\"\r\n",
        "Name;Note;Price\r\nAda;\"two\r\nlines\";1,5\r\n\"semi;colon\";\"say \"\"hi\"\"\";\r\n",
        "Name\tNote\r\nAda\t\"tab\there\"\r\n\t\r\n",
        "Name,Note\r\n",
        "Name\r\n\"\"\r\nx\r\n",
        ",Город,Город\r\nЁж 🍎,2024-01-15 meeting notes, padded \r\n",
        "\"\"\r\n",
    ];

    [TestMethod]
    public void ReadThenWriteGivesTheSameBytes()
    {
        foreach (var encoding in Enum.GetValues<CsvEncoding>())
        {
            foreach (var text in CanonicalFiles)
            {
                var bytes = Encode(text, encoding);

                var result = Read(bytes);
                var written = Write(result.Workbook.Sheets[0], result.Format);

                CollectionAssert.AreEqual(bytes, written, $"{encoding}: {text}");
            }
        }
    }

    [TestMethod]
    public void WriteThenReadGivesTheSameSheet()
    {
        const string Description =
            "Name||Name|Note¶Ada|x|1,5|\"quoted\"¶|;|\t|two\r\nlines¶ padded |=1+2|2024-01-15 meeting notes|semi;colon,comma\ttab¶|||";
        foreach (var separator in Enum.GetValues<CsvSeparator>())
        {
            foreach (var encoding in Enum.GetValues<CsvEncoding>())
            {
                var format = new CsvFormat(separator, encoding);

                var result = Read(Write(MakeSheet(Description), format));

                Assert.AreEqual(format, result.Format, $"{format}");
                Assert.AreEqual(Description, Describe(result.Workbook.Sheets[0]), $"{format}");
            }
        }
    }

    [TestMethod]
    public void WriteThenReadKeepsAHeaderOnlySheet()
    {
        foreach (var separator in Enum.GetValues<CsvSeparator>())
        {
            var format = new CsvFormat(separator, CsvEncoding.Utf8);

            var sheet = Read(Write(MakeSheet("a|b|c¶||"), format)).Workbook.Sheets[0];

            Assert.AreEqual("a|b|c¶||", Describe(sheet), $"{separator}");
        }
    }
}
