using System.Globalization;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Values;

namespace SpreadsheetEditor.Core.Tests;

// Cultures come from the platform's culture data (ICU on Linux). Without it (invariant globalization mode) GetCultureInfo throws and these
// tests fail; they are never skipped.
[TestClass]
public sealed class CellTextTests
{
    private static readonly CultureInfo English = CultureInfo.GetCultureInfo("en-US");
    private static readonly CultureInfo Russian = CultureInfo.GetCultureInfo("ru-RU");

    private static CultureInfo Culture(string name) => CultureInfo.GetCultureInfo(name);

    // The culture's own negative sign, so that the expected text does not depend on the culture data's choice of minus sign.
    private static string Negative(string name, string text) => Culture(name).NumberFormat.NegativeSign + text;

    [TestMethod]
    [DataRow("en-US", 3.0, "3")]
    [DataRow("ru-RU", 3.0, "3")]
    [DataRow("en-US", 3.5, "3.5")]
    [DataRow("ru-RU", 3.5, "3,5")]
    [DataRow("en-US", 0.1, "0.1")]
    [DataRow("ru-RU", 1234567.25, "1234567,25")]
    [DataRow("en-US", 0.00001, "0.00001")]
    [DataRow("en-US", 1e20, "100000000000000000000")]
    [DataRow("en-US", 1.5e-7, "0.00000015")]
    [DataRow("ru-RU", 0.30000000000000004, "0,30000000000000004")]
    [DataRow("en-US", -0.0, "0")]
    public void ShowsNumbers(string culture, double number, string text) =>
        Assert.AreEqual(text, CellText.Format(new CellValue(CellValueKind.Number, "", Number: number), Culture(culture)));

    [TestMethod]
    [DataRow("en-US")]
    [DataRow("ru-RU")]
    public void ShowsNegativeNumbersWithTheCultureSign(string culture)
    {
        string expected = Negative(culture, culture == "en-US" ? "2.25" : "2,25");
        Assert.AreEqual(expected, CellText.FormatNumber(-2.25, Culture(culture)));
    }

    [TestMethod]
    public void ShowsHugeNumbersWithAnExponent() =>
        Assert.AreEqual("1E+30", CellText.FormatNumber(1e30, English));

    [TestMethod]
    [DataRow("en-US")]
    [DataRow("ru-RU")]
    public void ShowsBooleansAsTrueAndFalse(string culture)
    {
        Assert.AreEqual("TRUE", CellText.Format(new CellValue(CellValueKind.Boolean, "", Boolean: true), Culture(culture)));
        Assert.AreEqual("FALSE", CellText.Format(new CellValue(CellValueKind.Boolean, "", Boolean: false), Culture(culture)));
    }

    [TestMethod]
    public void ShowsDatesInTheShortDatePattern()
    {
        var date = new CellValue(CellValueKind.Date, "", Date: new DateTime(2024, 1, 15));
        Assert.AreEqual("15.01.2024", CellText.Format(date, Russian));
        Assert.AreEqual(new DateTime(2024, 1, 15).ToString(English.DateTimeFormat.ShortDatePattern, English), CellText.Format(date, English));
        Assert.AreEqual("1/15/2024", CellText.Format(date, English));
    }

    [TestMethod]
    public void ShowsTheTimeWhenTheDateHasOne()
    {
        Assert.AreEqual("15.01.2024 13:30", CellText.FormatDate(new DateTime(2024, 1, 15, 13, 30, 0), Russian));
        Assert.AreEqual("15.01.2024 13:30:15", CellText.FormatDate(new DateTime(2024, 1, 15, 13, 30, 15), Russian));

        var date = new DateTime(2024, 1, 15, 13, 30, 0);
        string expected = "1/15/2024 " + date.ToString(English.DateTimeFormat.ShortTimePattern, English);
        Assert.AreEqual(expected, CellText.FormatDate(date, English));
    }

    [TestMethod]
    public void ShowsATimeOnlyAsATime()
    {
        Assert.AreEqual("13:30", CellText.FormatDate(DateSerial.TimeOnlyDay.AddHours(13.5), Russian));
        Assert.AreEqual("13:30:15", CellText.FormatDate(DateSerial.TimeOnlyDay + new TimeSpan(13, 30, 15), Russian));
        DateTime time = DateSerial.TimeOnlyDay.AddHours(13.5);
        Assert.AreEqual(time.ToString(English.DateTimeFormat.ShortTimePattern, English), CellText.FormatDate(time, English));
    }

    [TestMethod]
    [DataRow("en-US")]
    [DataRow("ru-RU")]
    public void ShowsFormulasAsTheirSource(string culture)
    {
        var formula = new CellValue(CellValueKind.Number, "", Number: 3, Formula: "SUM(A1:A2)");
        Assert.AreEqual("=SUM(A1:A2)", CellText.Format(formula, Culture(culture)));
    }

    [TestMethod]
    public void ShowsTextAndErrorsAsRead()
    {
        Assert.AreEqual("Hello", CellText.Format(new CellValue(CellValueKind.Text, "Hello"), English));
        Assert.AreEqual("#DIV/0!", CellText.Format(new CellValue(CellValueKind.Error, "#DIV/0!"), Russian));
    }

    [TestMethod]
    [DataRow("en-US", "3", 3.0)]
    [DataRow("en-US", "0", 0.0)]
    [DataRow("en-US", "0.5", 0.5)]
    [DataRow("en-US", "1234567.25", 1234567.25)]
    [DataRow("en-US", "1.50", 1.5)]
    [DataRow("ru-RU", "3,5", 3.5)]
    [DataRow("ru-RU", "0,001", 0.001)]
    [DataRow("ru-RU", "123456789012345", 123456789012345.0)]
    public void ParsesNumbers(string culture, string text, double number)
    {
        CellValue value = CellText.Parse(text, Culture(culture));
        Assert.AreEqual(CellValueKind.Number, value.Kind);
        Assert.AreEqual<double?>(number, value.Number);
        Assert.AreEqual(text, value.DisplayText);
    }

    [TestMethod]
    [DataRow("en-US")]
    [DataRow("ru-RU")]
    public void ParsesNegativeNumbersWithTheCultureSign(string culture)
    {
        CellValue value = CellText.Parse(Negative(culture, "42"), Culture(culture));
        Assert.AreEqual(CellValueKind.Number, value.Kind);
        Assert.AreEqual<double?>(-42.0, value.Number);
    }

    [TestMethod]
    [DataRow("en-US", "TRUE", true)]
    [DataRow("en-US", "true", true)]
    [DataRow("ru-RU", "False", false)]
    [DataRow("ru-RU", "FALSE", false)]
    public void ParsesBooleans(string culture, string text, bool boolean)
    {
        CellValue value = CellText.Parse(text, Culture(culture));
        Assert.AreEqual(CellValueKind.Boolean, value.Kind);
        Assert.AreEqual<bool?>(boolean, value.Boolean);
        Assert.AreEqual(text, value.DisplayText);
    }

    [TestMethod]
    [DataRow("en-US")]
    [DataRow("ru-RU")]
    public void ParsesFormulas(string culture)
    {
        CellValue value = CellText.Parse("=A1+1", Culture(culture));
        Assert.AreEqual("A1+1", value.Formula);
        Assert.AreEqual("=A1+1", value.DisplayText);
        Assert.IsNull(value.Number);
    }

    [TestMethod]
    public void ParsesDates()
    {
        Assert.AreEqual(new DateTime(2024, 1, 15), ParseDate("15.01.2024", Russian));
        Assert.AreEqual(new DateTime(2024, 1, 15, 13, 30, 0), ParseDate("15.01.2024 13:30", Russian));
        Assert.AreEqual(new DateTime(2024, 1, 15, 13, 30, 15), ParseDate("15.01.2024 13:30:15", Russian));
        Assert.AreEqual(new DateTime(2024, 1, 15), ParseDate("1/15/2024", English));
        Assert.AreEqual(new DateTime(2024, 1, 15, 13, 30, 0), ParseDate("1/15/2024 1:30 PM", English));
        Assert.AreEqual(new DateTime(2024, 1, 15, 13, 30, 15), ParseDate("1/15/2024 1:30:15 PM", English));
    }

    [TestMethod]
    public void ParsesTimes()
    {
        Assert.AreEqual(DateSerial.TimeOnlyDay.AddHours(13.5), ParseDate("13:30", Russian));
        Assert.AreEqual(DateSerial.TimeOnlyDay + new TimeSpan(13, 30, 15), ParseDate("13:30:15", Russian));
        Assert.AreEqual(DateSerial.TimeOnlyDay.AddHours(13.5), ParseDate("1:30 PM", English));
    }

    [TestMethod]
    [DataRow("en-US", "2024-01-15 meeting notes")]
    [DataRow("ru-RU", "2024-01-15 meeting notes")]
    [DataRow("en-US", "2024-01-15")]
    [DataRow("ru-RU", "1/15/2024")]
    [DataRow("en-US", "15.01.2024")]
    [DataRow("ru-RU", "31.02.2024")]
    [DataRow("en-US", "1/15/1899")]
    [DataRow("en-US", "1,5")]
    [DataRow("en-US", "1,000")]
    [DataRow("ru-RU", "3.5")]
    [DataRow("ru-RU", "1 000")]
    [DataRow("en-US", "007")]
    [DataRow("en-US", "00123")]
    [DataRow("en-US", " 3")]
    [DataRow("en-US", "3 ")]
    [DataRow("en-US", "+5")]
    [DataRow("en-US", "1E5")]
    [DataRow("en-US", "1e-3")]
    [DataRow("en-US", "3.")]
    [DataRow("en-US", ".5")]
    [DataRow("en-US", "$5")]
    [DataRow("en-US", "50%")]
    [DataRow("en-US", "NaN")]
    [DataRow("en-US", "Infinity")]
    [DataRow("en-US", "1234567890123456")]
    [DataRow("en-US", "4111111111111111")]
    [DataRow("en-US", " TRUE")]
    [DataRow("en-US", "yes")]
    [DataRow("ru-RU", "ИСТИНА")]
    [DataRow("en-US", "=")]
    [DataRow("en-US", "#DIV/0!")]
    [DataRow("en-US", "Hello, world")]
    [DataRow("ru-RU", "Привет")]
    public void KeepsOtherTextExactlyAsTyped(string culture, string text)
    {
        CellValue value = CellText.Parse(text, Culture(culture));
        Assert.AreEqual(CellValueKind.Text, value.Kind);
        Assert.IsNull(value.Formula);
        Assert.AreEqual(text, value.DisplayText);
        Assert.AreEqual(text, CellText.Format(value, Culture(culture)));
    }

    [TestMethod]
    [DataRow("en-US")]
    [DataRow("ru-RU")]
    public void ValuesRoundTripThroughTheirText(string culture)
    {
        CultureInfo info = Culture(culture);
        CellValue[] values =
        [
            new(CellValueKind.Number, "", Number: 3),
            new(CellValueKind.Number, "", Number: -3.25),
            new(CellValueKind.Number, "", Number: 0.1),
            new(CellValueKind.Number, "", Number: 1234567.875),
            new(CellValueKind.Number, "", Number: 0.00001),
            new(CellValueKind.Number, "", Number: 1e14),
            new(CellValueKind.Boolean, "", Boolean: true),
            new(CellValueKind.Boolean, "", Boolean: false),
            new(CellValueKind.Date, "", Date: new DateTime(2024, 1, 15)),
            new(CellValueKind.Date, "", Date: new DateTime(2024, 12, 31, 9, 5, 0)),
            new(CellValueKind.Date, "", Date: new DateTime(1900, 3, 1, 23, 59, 59)),
            new(CellValueKind.Date, "", Date: DateSerial.TimeOnlyDay.AddHours(7.25)),
            new(CellValueKind.Date, "", Date: DateSerial.TimeOnlyDay + new TimeSpan(18, 0, 30)),
            new(CellValueKind.Text, "", Formula: "SUM(A1:A3)*2"),
            new(CellValueKind.Text, "Plain text"),
        ];

        foreach (CellValue value in values)
        {
            string text = CellText.Format(value, info);
            CellValue back = CellText.Parse(text, info);
            string label = culture + ": " + text;
            Assert.AreEqual(value.Kind, back.Kind, label);
            Assert.AreEqual(value.Number, back.Number, label);
            Assert.AreEqual(value.Boolean, back.Boolean, label);
            Assert.AreEqual(value.Date, back.Date, label);
            Assert.AreEqual(value.Formula, back.Formula, label);
            Assert.AreEqual(text, back.DisplayText, label);
            Assert.AreEqual(text, CellText.Format(back, info), label);
        }
    }

    private static DateTime ParseDate(string text, CultureInfo culture)
    {
        CellValue value = CellText.Parse(text, culture);
        Assert.AreEqual(CellValueKind.Date, value.Kind, text);
        Assert.IsNotNull(value.Date, text);
        return value.Date.GetValueOrDefault();
    }
}
