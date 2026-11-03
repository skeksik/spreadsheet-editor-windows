using System.Globalization;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Values;

namespace SpreadsheetEditor.Core.Tests;

[TestClass]
public sealed class EntryValueTests
{
    private static readonly CultureInfo Russian = CultureInfo.GetCultureInfo("ru-RU");
    private static readonly NumberFormat IsoDate = new(164, "yyyy-mm-dd");

    private static readonly CellValue DateCell =
        new(CellValueKind.Date, "15.01.2024", Number: 45306, Date: new DateTime(2024, 1, 15), Format: IsoDate);

    [TestMethod]
    public void AnUneditedValueIsNeverParsedAgain()
    {
        // "007" would stay text and "TRUE" would be a boolean if parsed; unedited, the value as read is kept.
        var number = new CellValue(CellValueKind.Number, "007", Number: 7, Format: new NumberFormat(165, "000"));
        Assert.AreSame(number, EntryValue.Resolve(new Entry { Text = "007", Original = number }, Russian));

        var text = new CellValue(CellValueKind.Text, "TRUE");
        Assert.AreSame(text, EntryValue.Resolve(new Entry { Text = "TRUE", Original = text }, Russian));

        Assert.AreSame(DateCell, EntryValue.Resolve(new Entry { Text = "15.01.2024", Original = DateCell }, Russian));
    }

    [TestMethod]
    public void EmptyTextClearsTheCell()
    {
        Assert.IsNull(EntryValue.Resolve(new Entry { Text = "", Original = DateCell }, Russian));
        Assert.IsNull(EntryValue.Resolve(new Entry(), Russian));
    }

    [TestMethod]
    public void EditedTextIsParsed()
    {
        CellValue? value = EntryValue.Resolve(new Entry { Text = "3,5", Original = new CellValue(CellValueKind.Text, "x") }, Russian);
        Assert.IsNotNull(value);
        Assert.AreEqual(CellValueKind.Number, value.Kind);
        Assert.AreEqual<double?>(3.5, value.Number);
        Assert.IsNull(value.Format);

        CellValue? fresh = EntryValue.Resolve(new Entry { Text = "Привет" }, Russian);
        Assert.IsNotNull(fresh);
        Assert.AreEqual(CellValueKind.Text, fresh.Kind);
        Assert.AreEqual("Привет", fresh.DisplayText);
    }

    [TestMethod]
    public void AnEditedDateKeepsTheCellsDateFormat()
    {
        CellValue? value = EntryValue.Resolve(new Entry { Text = "16.01.2024", Original = DateCell }, Russian);
        Assert.IsNotNull(value);
        Assert.AreEqual(CellValueKind.Date, value.Kind);
        Assert.AreEqual<DateTime?>(new DateTime(2024, 1, 16), value.Date);
        Assert.IsNull(value.Number);
        Assert.AreEqual(IsoDate, value.Format);
    }

    [TestMethod]
    [DataRow("16.01.2024", 14)]
    [DataRow("16.01.2024 13:30", 22)]
    [DataRow("13:30", 20)]
    [DataRow("13:30:15", 21)]
    public void ADateTypedIntoAnotherCellGetsTheDefaultFormat(string text, int id)
    {
        CellValue? value = EntryValue.Resolve(new Entry { Text = text, Original = new CellValue(CellValueKind.Number, "1", Number: 1) }, Russian);
        Assert.IsNotNull(value);
        Assert.AreEqual(CellValueKind.Date, value.Kind);
        Assert.AreEqual(new NumberFormat(id), value.Format);
    }

    [TestMethod]
    public void ANumberTypedIntoADateCellIsShownAsANumber()
    {
        CellValue? number = EntryValue.Resolve(new Entry { Text = "42", Original = DateCell }, Russian);
        Assert.IsNotNull(number);
        Assert.AreEqual(CellValueKind.Number, number.Kind);
        Assert.AreEqual(NumberFormat.General, number.Format);

        CellValue? text = EntryValue.Resolve(new Entry { Text = "soon", Original = DateCell }, Russian);
        Assert.IsNotNull(text);
        Assert.AreEqual(CellValueKind.Text, text.Kind);
        Assert.IsNull(text.Format);
    }

    [TestMethod]
    [DataRow("1/2")]
    [DataRow("15.01.2024")]
    [DataRow("007")]
    [DataRow("42")]
    [DataRow("TRUE")]
    [DataRow("=A1+1")]
    public void ACellFormattedAsTextKeepsWhatIsTyped(string typed)
    {
        foreach (NumberFormat format in new[] { new NumberFormat(49), new NumberFormat(164, "@") })
        {
            var original = new CellValue(CellValueKind.Text, "old", Format: format);
            CellValue? value = EntryValue.Resolve(new Entry { Text = typed, Original = original }, Russian);
            Assert.IsNotNull(value);
            Assert.AreEqual(CellValueKind.Text, value.Kind);
            Assert.AreEqual(typed, value.DisplayText);
            Assert.IsNull(value.Formula);
            Assert.IsNull(value.Format);
        }
    }

    [TestMethod]
    public void RecognisesTheTextFormat()
    {
        Assert.IsTrue(NumberFormats.IsText(new NumberFormat(49)));
        Assert.IsTrue(NumberFormats.IsText(new NumberFormat(164, "@")));
        Assert.IsTrue(NumberFormats.IsText(NumberFormats.Text));
        Assert.IsFalse(NumberFormats.IsText(new NumberFormat(0)));
        Assert.IsFalse(NumberFormats.IsText(new NumberFormat(49, "0.00")));
        Assert.IsFalse(NumberFormats.IsText(new NumberFormat(164, "0;-0;0;@")));
        Assert.IsFalse(NumberFormats.IsText(null));
    }
}
