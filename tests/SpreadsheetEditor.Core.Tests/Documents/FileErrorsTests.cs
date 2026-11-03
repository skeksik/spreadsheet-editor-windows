using SpreadsheetEditor.Core.Csv;
using SpreadsheetEditor.Core.Documents;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Documents;

[TestClass]
public sealed class FileErrorsTests
{
    [TestMethod]
    [DataRow(SpreadsheetReadError.NotFound, FileProblem.NotFound)]
    [DataRow(SpreadsheetReadError.AccessDenied, FileProblem.AccessDenied)]
    [DataRow(SpreadsheetReadError.ReadFailed, FileProblem.ReadFailed)]
    [DataRow(SpreadsheetReadError.NotAnXlsx, FileProblem.NotAWorkbook)]
    [DataRow(SpreadsheetReadError.LegacyOrEncrypted, FileProblem.LegacyOrEncrypted)]
    [DataRow(SpreadsheetReadError.Corrupt, FileProblem.Damaged)]
    public void ClassifiesReadErrors(SpreadsheetReadError reason, FileProblem expected) =>
        Assert.AreEqual(expected, FileErrors.Classify(new SpreadsheetReadException(reason, "detail")));

    [TestMethod]
    [DataRow(CsvReadFailure.Empty, FileProblem.CsvEmpty)]
    [DataRow(CsvReadFailure.UnsupportedEncoding, FileProblem.CsvUnsupportedEncoding)]
    [DataRow(CsvReadFailure.UnterminatedQuote, FileProblem.CsvUnterminatedQuote)]
    public void ClassifiesCsvErrors(CsvReadFailure reason, FileProblem expected) =>
        Assert.AreEqual(expected, FileErrors.Classify(new CsvReadException(reason, 1, "detail")));

    [TestMethod]
    [DataRow(SpreadsheetWriteError.NotSupported, FileProblem.SaveNotSupported)]
    [DataRow(SpreadsheetWriteError.InvalidWorkbook, FileProblem.InvalidDocument)]
    [DataRow(SpreadsheetWriteError.SourceDamaged, FileProblem.SourceDamaged)]
    [DataRow(SpreadsheetWriteError.AccessDenied, FileProblem.AccessDenied)]
    [DataRow(SpreadsheetWriteError.WriteFailed, FileProblem.WriteFailed)]
    public void ClassifiesWriteErrors(SpreadsheetWriteError reason, FileProblem expected) =>
        Assert.AreEqual(expected, FileErrors.Classify(new SpreadsheetWriteException(reason, "detail")));

    [TestMethod]
    [DataRow(unchecked((int)0x80070020))]
    [DataRow(unchecked((int)0x80070021))]
    public void AFileHeldByAnotherProgramIsInUse(int hresult)
    {
        var inner = new IOException("The process cannot access the file because it is being used by another process.", hresult);

        Assert.AreEqual(FileProblem.InUse, FileErrors.Classify(new SpreadsheetWriteException(SpreadsheetWriteError.WriteFailed, "detail", inner)));
        Assert.AreEqual(FileProblem.InUse, FileErrors.Classify(new SpreadsheetReadException(SpreadsheetReadError.ReadFailed, "detail", inner)));
    }

    [TestMethod]
    public void AnotherIoErrorKeepsItsReason()
    {
        var inner = new IOException("The disk is full.", unchecked((int)0x80070070));

        Assert.AreEqual(FileProblem.WriteFailed, FileErrors.Classify(new SpreadsheetWriteException(SpreadsheetWriteError.WriteFailed, "detail", inner)));
        Assert.AreEqual(FileProblem.Unknown, FileErrors.Classify(new InvalidOperationException("detail", new IOException("x", unchecked((int)0x80070020)))));
    }

    [TestMethod]
    public void EveryReasonHasAProblem()
    {
        foreach (var reason in Enum.GetValues<SpreadsheetReadError>())
        {
            Assert.AreNotEqual(FileProblem.Unknown, FileErrors.Classify(new SpreadsheetReadException(reason, "detail")), reason.ToString());
        }
        foreach (var reason in Enum.GetValues<CsvReadFailure>())
        {
            Assert.AreNotEqual(FileProblem.Unknown, FileErrors.Classify(new CsvReadException(reason, 1, "detail")), reason.ToString());
        }
        foreach (var reason in Enum.GetValues<SpreadsheetWriteError>())
        {
            Assert.AreNotEqual(FileProblem.Unknown, FileErrors.Classify(new SpreadsheetWriteException(reason, "detail")), reason.ToString());
        }
    }

    [TestMethod]
    public void AnyOtherExceptionIsUnknown()
    {
        Assert.AreEqual(FileProblem.Unknown, FileErrors.Classify(new InvalidOperationException("detail")));
        Assert.AreEqual(FileProblem.Unknown, FileErrors.Classify(new IOException("detail")));
    }
}
