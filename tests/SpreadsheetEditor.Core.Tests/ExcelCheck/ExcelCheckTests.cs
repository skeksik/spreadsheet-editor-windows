using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.ExcelCheck;

/// <summary>
/// Writes the Excel check's sample workbooks (<see cref="ExcelCheckSamples"/>, docs/notes/excel-check.md) with the app's save path,
/// validates each package against the Open XML schemas, reads it back and checks what the notes tell the owner to look for. When the
/// environment variable <c>EXCEL_CHECK_DIR</c> names a folder, the files are also copied there, for opening in Excel.
/// </summary>
[TestClass]
public sealed partial class ExcelCheckTests
{
    /// <summary>The environment variable naming the folder that receives a copy of each sample.</summary>
    public const string FolderVariable = "EXCEL_CHECK_DIR";

    /// <summary>The control file: a source workbook as the test builders make it, which the app does not touch.</summary>
    public const string ControlFileName = "00-control.xlsx";

    private string folder = "";

    public TestContext TestContext { get; set; } = null!;

    [TestInitialize]
    public void CreateFolder()
    {
        folder = Path.Combine(Path.GetTempPath(), "excel-check-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
    }

    [TestCleanup]
    public void DeleteFolder() => Directory.Delete(folder, recursive: true);

    [TestMethod]
    public void EverySampleHasATest()
    {
        var tested = typeof(ExcelCheckTests).GetMethods()
            .Where(m => m.GetCustomAttributes(typeof(TestMethodAttribute), false).Length > 0)
            .Select(m => m.Name)
            .ToHashSet();
        foreach (var (fileName, _) in ExcelCheckSamples.All)
        {
            var words = Path.GetFileNameWithoutExtension(fileName)[3..].Split('-');
            var name = "Sample" + string.Concat(words.Select(w => char.ToUpperInvariant(w[0]) + w[1..]));
            Assert.Contains(name, tested, $"No test for {fileName}.");
        }
    }

    /// <summary>
    /// Saves the sample with <see cref="XlsxWriter.Write(Workbook, string, System.Globalization.CultureInfo, CancellationToken)"/>, checks
    /// that a second save gives the same file, validates the package, copies it to <see cref="FolderVariable"/>'s folder when set, and
    /// returns the saved file read back.
    /// </summary>
    private Saved Save(string fileName)
    {
        var build = ExcelCheckSamples.All.Single(s => s.FileName == fileName).Build;
        var workbook = build();
        var path = Path.Combine(folder, fileName);
        XlsxWriter.Write(workbook, path, ExcelCheckSamples.Culture);
        var bytes = File.ReadAllBytes(path);
        XlsxWriter.Write(workbook, path, ExcelCheckSamples.Culture);
        CollectionAssert.AreEqual(bytes, File.ReadAllBytes(path), "Two saves differ.");

        AssertValid(bytes, fileName);

        Copy(path);
        return new Saved(bytes, XlsxReader.Read(path, ExcelCheckSamples.Culture));
    }

    /// <summary>Copies a written file to <see cref="FolderVariable"/>'s folder when the variable is set.</summary>
    private void Copy(string path)
    {
        var target = Environment.GetEnvironmentVariable(FolderVariable);
        if (!string.IsNullOrWhiteSpace(target))
        {
            Directory.CreateDirectory(target);
            File.Copy(path, Path.Combine(target, Path.GetFileName(path)), overwrite: true);
            TestContext.WriteLine($"Copied {Path.GetFileName(path)} to {target}.");
        }
    }

    private static void AssertValid(byte[] bytes, string fileName)
    {
        var errors = PackageValidation.Errors(bytes);
        Assert.IsEmpty(errors, $"{fileName} is not valid against the Open XML schemas ({PackageValidation.Version}):\n" + string.Join("\n", errors));
    }

    private sealed class Saved(byte[] bytes, Workbook workbook)
    {
        public byte[] Bytes { get; } = bytes;

        public Workbook Workbook { get; } = workbook;

        public Sheet Sheet(string name) => Workbook.Sheets.Single(s => s.Name == name);
    }
}
