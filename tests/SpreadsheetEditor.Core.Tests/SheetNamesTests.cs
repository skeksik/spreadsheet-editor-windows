using System.Globalization;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests;

[TestClass]
public sealed class SheetNamesTests
{
    private static readonly string[] NoNames = [];

    private static List<string> Names(params string[] names) => [.. names];

    // An emoji outside the Basic Multilingual Plane: one code point, two UTF-16 code units.
    private const string Smile = "\U0001F600";

    [TestMethod]
    [DataRow("A")]
    [DataRow("Sheet1")]
    [DataRow("Sam's orders")]
    [DataRow("Лист 1")]
    [DataRow("a'b")]
    [DataRow("(Copy)")]
    [DataRow(" ")]
    public void AcceptsValidNames(string name) => Assert.AreEqual(SheetNameCheck.Valid, SheetNames.Check(name, NoNames));

    [TestMethod]
    public void RejectsEmptyName() => Assert.AreEqual(SheetNameError.Empty, SheetNames.Check("", NoNames).Error);

    [TestMethod]
    public void AcceptsMaximumLength()
    {
        var name = new string('a', SheetNames.MaxLength);
        Assert.IsTrue(SheetNames.Check(name, NoNames).IsValid);
    }

    [TestMethod]
    public void RejectsOneOverMaximumLength()
    {
        var name = new string('a', SheetNames.MaxLength + 1);
        Assert.AreEqual(SheetNameError.TooLong, SheetNames.Check(name, NoNames).Error);
    }

    [TestMethod]
    public void CountsLengthInUtf16CodeUnits()
    {
        var fits = new string('a', SheetNames.MaxLength - 2) + Smile;
        var over = new string('a', SheetNames.MaxLength - 1) + Smile;
        Assert.IsTrue(SheetNames.Check(fits, NoNames).IsValid);
        Assert.AreEqual(SheetNameError.TooLong, SheetNames.Check(over, NoNames).Error);
    }

    [TestMethod]
    [DataRow(':')]
    [DataRow('\\')]
    [DataRow('/')]
    [DataRow('?')]
    [DataRow('*')]
    [DataRow('[')]
    [DataRow(']')]
    [DataRow('\u0000')]
    [DataRow('\u0003')]
    public void RejectsInvalidCharacterAndNamesIt(char character)
    {
        var check = SheetNames.Check($"a{character}b", NoNames);
        Assert.AreEqual(new SheetNameCheck(SheetNameError.InvalidCharacter, character), check);
        Assert.IsFalse(check.IsValid);
    }

    [TestMethod]
    public void NamesTheFirstInvalidCharacter() =>
        Assert.AreEqual(new SheetNameCheck(SheetNameError.InvalidCharacter, '?'), SheetNames.Check("a?b*c", NoNames));

    [TestMethod]
    public void RejectsLeadingApostrophe() =>
        Assert.AreEqual(SheetNameError.LeadingApostrophe, SheetNames.Check("'a", NoNames).Error);

    [TestMethod]
    public void RejectsTrailingApostrophe() =>
        Assert.AreEqual(SheetNameError.TrailingApostrophe, SheetNames.Check("a'", NoNames).Error);

    [TestMethod]
    [DataRow("'")]
    [DataRow("''")]
    public void RejectsApostropheOnlyNamesAsLeading(string name) =>
        Assert.AreEqual(SheetNameError.LeadingApostrophe, SheetNames.Check(name, NoNames).Error);

    [TestMethod]
    [DataRow("History")]
    [DataRow("history")]
    [DataRow("HISTORY")]
    public void RejectsReservedName(string name) =>
        Assert.AreEqual(SheetNameError.Reserved, SheetNames.Check(name, NoNames).Error);

    [TestMethod]
    [DataRow("History1")]
    [DataRow("My History")]
    public void AcceptsNamesContainingReservedWord(string name) => Assert.IsTrue(SheetNames.Check(name, NoNames).IsValid);

    [TestMethod]
    [DataRow("Data", "Data")]
    [DataRow("Data", "DATA")]
    [DataRow("data", "Data")]
    [DataRow("Лист", "ЛИСТ")]
    [DataRow("Ёлка", "ёлка")]
    [DataRow("Résumé", "RÉSUMÉ")]
    [DataRow("Σοφία", "ΣΟΦΊΑ")]
    public void RejectsDuplicatesIgnoringCase(string existing, string name) =>
        Assert.AreEqual(SheetNameError.Duplicate, SheetNames.Check(name, Names("Other", existing)).Error);

    [TestMethod]
    [DataRow("Resume", "Résumé")]
    [DataRow("Data", "Data ")]
    [DataRow("Лист", "Лист1")]
    public void AcceptsNamesThatDifferBeyondCase(string existing, string name) =>
        Assert.IsTrue(SheetNames.Check(name, Names(existing)).IsValid);

    [TestMethod]
    public void RenameMayKeepTheNameOrChangeItsCase()
    {
        string[] names = ["Data", "Other"];
        Assert.IsTrue(SheetNames.Check("Data", names, renamedName: "Data").IsValid);
        Assert.IsTrue(SheetNames.Check("DATA", names, renamedName: "Data").IsValid);
    }

    [TestMethod]
    public void RenameStillRejectsOtherSheetsNames() =>
        Assert.AreEqual(SheetNameError.Duplicate, SheetNames.Check("other", Names("Data", "Other"), renamedName: "Data").Error);

    [TestMethod]
    public void RenameSkipsOnlyOneSheet() =>
        Assert.AreEqual(SheetNameError.Duplicate, SheetNames.Check("Data", Names("Data", "Data"), renamedName: "Data").Error);

    [TestMethod]
    public void RulesComeBeforeDuplicates() =>
        Assert.AreEqual(SheetNameError.Reserved, SheetNames.Check("History", Names("History")).Error);

    [TestMethod]
    public void RejectsNullArguments()
    {
        Assert.ThrowsExactly<ArgumentNullException>(() => SheetNames.Check(null!, NoNames));
        Assert.ThrowsExactly<ArgumentNullException>(() => SheetNames.Check("a", null!));
        Assert.ThrowsExactly<ArgumentNullException>(() => SheetNames.NewSheetName(null!, NoNames));
        Assert.ThrowsExactly<ArgumentNullException>(() => SheetNames.CopyName(null!, NoNames));
        Assert.ThrowsExactly<ArgumentNullException>(() => SheetNames.CopyName("a", null!));
    }

    [TestMethod]
    public void NewSheetNameStartsAtOne() => Assert.AreEqual("Sheet1", SheetNames.NewSheetName("Sheet", NoNames));

    [TestMethod]
    public void NewSheetNameTakesTheSmallestFreeNumberIgnoringCase() =>
        Assert.AreEqual("Sheet2", SheetNames.NewSheetName("Sheet", Names("SHEET1", "sheet3", "Other")));

    [TestMethod]
    public void NewSheetNameUsesTheLocalisedBase() =>
        Assert.AreEqual("Лист3", SheetNames.NewSheetName("Лист", Names("лист1", "ЛИСТ2")));

    [TestMethod]
    public void NewSheetNameShortensALongBase()
    {
        var baseName = new string('b', 40);
        var first = SheetNames.NewSheetName(baseName, NoNames);
        Assert.AreEqual(new string('b', 30) + "1", first);

        var names = Enumerable.Range(1, 9).Select(n => new string('b', 30) + n).ToArray();
        Assert.AreEqual(new string('b', 29) + "10", SheetNames.NewSheetName(baseName, names));
    }

    [TestMethod]
    public void NewSheetNameNeverSplitsASurrogatePair()
    {
        var baseName = new string('b', 29) + Smile + "c";
        var name = SheetNames.NewSheetName(baseName, NoNames);
        Assert.AreEqual(new string('b', 29) + "1", name);
        Assert.IsTrue(SheetNames.Check(name, NoNames).IsValid);
    }

    [TestMethod]
    public void CopyNameAddsCopy() => Assert.AreEqual("Data (Copy)", SheetNames.CopyName("Data", Names("Data")));

    [TestMethod]
    public void CopyNameNumbersWhenTaken()
    {
        Assert.AreEqual("Data (Copy 1)", SheetNames.CopyName("Data", Names("Data", "Data (Copy)")));
        Assert.AreEqual("Data (Copy 2)", SheetNames.CopyName("Data", Names("Data", "data (copy)", "DATA (COPY 1)")));
        Assert.AreEqual("Data (Copy 1)", SheetNames.CopyName("Data", Names("Data", "Data (Copy)", "Data (Copy 2)")));
    }

    [TestMethod]
    public void CopyNameOfACopyAddsAnotherSuffix() =>
        Assert.AreEqual("Data (Copy) (Copy)", SheetNames.CopyName("Data (Copy)", Names("Data", "Data (Copy)")));

    [TestMethod]
    public void CopyNameKeepsTheSuffixUntranslated() =>
        Assert.AreEqual("Лист1 (Copy)", SheetNames.CopyName("Лист1", Names("Лист1")));

    [TestMethod]
    public void CopyNameFitsExactlyWithoutShortening()
    {
        var source = new string('a', SheetNames.MaxLength - " (Copy)".Length);
        var name = SheetNames.CopyName(source, Names(source));
        Assert.AreEqual(source + " (Copy)", name);
        Assert.AreEqual(SheetNames.MaxLength, name.Length);
    }

    [TestMethod]
    public void CopyNameShortensTheBaseToFit()
    {
        var source = new string('a', SheetNames.MaxLength);
        var first = SheetNames.CopyName(source, Names(source));
        Assert.AreEqual(new string('a', 24) + " (Copy)", first);

        var second = SheetNames.CopyName(source, Names(source, first));
        Assert.AreEqual(new string('a', 22) + " (Copy 1)", second);

        var names = new List<string> { source, first };
        names.AddRange(Enumerable.Range(1, 9).Select(n => new string('a', 22) + " (Copy " + n.ToString(CultureInfo.InvariantCulture) + ")"));
        var tenth = SheetNames.CopyName(source, names);
        Assert.AreEqual(new string('a', 21) + " (Copy 10)", tenth);
        Assert.AreEqual(SheetNames.MaxLength, tenth.Length);
    }

    [TestMethod]
    public void CopyNameNeverSplitsASurrogatePair()
    {
        // The cut for " (Copy)" falls after 24 code units, between the two halves of the emoji.
        var source = new string('a', 23) + Smile + "bbbb";
        var name = SheetNames.CopyName(source, Names(source));
        Assert.AreEqual(new string('a', 23) + " (Copy)", name);
        Assert.IsTrue(SheetNames.Check(name, Names(source)).IsValid);
    }

    [TestMethod]
    public void CopyNameKeepsAWholeSurrogatePairThatFits()
    {
        var source = new string('a', 22) + Smile + "bbbb";
        Assert.AreEqual(new string('a', 22) + Smile + " (Copy)", SheetNames.CopyName(source, Names(source)));
    }

    [TestMethod]
    public void GeneratedNamesPassTheRules()
    {
        string[] names = ["Sheet1", "Data", "Data (Copy)"];
        Assert.IsTrue(SheetNames.Check(SheetNames.NewSheetName("Sheet", names), names).IsValid);
        Assert.IsTrue(SheetNames.Check(SheetNames.CopyName("Data", names), names).IsValid);
    }
}
