using SpreadsheetEditor.Core.Documents;

namespace SpreadsheetEditor.Core.Tests.Documents;

[TestClass]
public sealed class WindowTitlesTests
{
    private const string WithFile = "{0} - Spreadsheet Editor";
    private const string App = "Spreadsheet Editor";

    [TestMethod]
    public void IsTheAppNameWithNoDocument() => Assert.AreEqual(App, WindowTitles.Compose(null, modified: true, WithFile, App));

    [TestMethod]
    public void ShowsTheFileName() => Assert.AreEqual("plants.xlsx - Spreadsheet Editor", WindowTitles.Compose("plants.xlsx", false, WithFile, App));

    [TestMethod]
    public void MarksUnsavedChangesBeforeTheFileName() =>
        Assert.AreEqual("*plants.xlsx - Spreadsheet Editor", WindowTitles.Compose("plants.xlsx", true, WithFile, App));

    [TestMethod]
    public void KeepsBracesInTheFileName() =>
        Assert.AreEqual("{0}.csv - Spreadsheet Editor", WindowTitles.Compose("{0}.csv", false, WithFile, App));
}
