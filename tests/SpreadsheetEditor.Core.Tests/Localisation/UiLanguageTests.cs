using System.Globalization;
using SpreadsheetEditor.Core.Localisation;

namespace SpreadsheetEditor.Core.Tests.Localisation;

[TestClass]
public sealed class UiLanguageTests
{
    [TestMethod]
    [DataRow("en-US", UiLanguage.English)]
    [DataRow("en-GB", UiLanguage.English)]
    [DataRow("en", UiLanguage.English)]
    [DataRow("ru-RU", UiLanguage.Russian)]
    [DataRow("ru-UA", UiLanguage.Russian)]
    [DataRow("ru", UiLanguage.Russian)]
    [DataRow("de-DE", UiLanguage.English)]
    [DataRow("uk-UA", UiLanguage.English)]
    [DataRow("", UiLanguage.English)]
    public void SystemFollowsWindowsWhenTheAppHasItsLanguageOtherwiseEnglish(string windows, string expected) =>
        Assert.AreEqual(expected, UiLanguage.Resolve(LanguagePreference.System, CultureInfo.GetCultureInfo(windows)));

    [TestMethod]
    [DataRow("en-US")]
    [DataRow("ru-RU")]
    [DataRow("de-DE")]
    public void AnOverrideWinsOverWindows(string windows)
    {
        CultureInfo culture = CultureInfo.GetCultureInfo(windows);
        Assert.AreEqual(UiLanguage.English, UiLanguage.Resolve(LanguagePreference.English, culture));
        Assert.AreEqual(UiLanguage.Russian, UiLanguage.Resolve(LanguagePreference.Russian, culture));
    }

    [TestMethod]
    [DataRow("en", LanguagePreference.English)]
    [DataRow(" RU ", LanguagePreference.Russian)]
    [DataRow("system", LanguagePreference.System)]
    [DataRow(null, LanguagePreference.System)]
    [DataRow("", LanguagePreference.System)]
    [DataRow("de", LanguagePreference.System)]
    [DataRow("en-US", LanguagePreference.System)]
    public void ParsesStoredValuesAndTreatsUnknownOnesAsSystem(string? stored, LanguagePreference expected) =>
        Assert.AreEqual(expected, UiLanguage.ParsePreference(stored));

    [TestMethod]
    [DataRow(LanguagePreference.System)]
    [DataRow(LanguagePreference.English)]
    [DataRow(LanguagePreference.Russian)]
    public void StoredValuesReadBack(LanguagePreference preference) =>
        Assert.AreEqual(preference, UiLanguage.ParsePreference(UiLanguage.ToStoredValue(preference)));
}
