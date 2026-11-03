using SpreadsheetEditor.Core.Values;

namespace SpreadsheetEditor.Core.Tests;

[TestClass]
public sealed class DateSerialTests
{
    [TestMethod]
    [DataRow(1.0, 1900, 1, 1)]
    [DataRow(59.0, 1900, 2, 28)]
    [DataRow(61.0, 1900, 3, 1)]
    [DataRow(39448.0, 2008, 1, 1)]
    [DataRow(40729.0, 2011, 7, 5)]
    [DataRow(2958465.0, 9999, 12, 31)]
    public void Reads1900Serials(double serial, int year, int month, int day)
    {
        Assert.IsTrue(DateSerial.TryToDateTime(serial, DateSystem.Date1900, out DateTime value));
        Assert.AreEqual(new DateTime(year, month, day), value);
    }

    [TestMethod]
    [DataRow(1.0, 1904, 1, 2)]
    [DataRow(39267.0, 2011, 7, 5)]
    [DataRow(2957003.0, 9999, 12, 31)]
    public void Reads1904Serials(double serial, int year, int month, int day)
    {
        Assert.IsTrue(DateSerial.TryToDateTime(serial, DateSystem.Date1904, out DateTime value));
        Assert.AreEqual(new DateTime(year, month, day), value);
    }

    [TestMethod]
    public void Serial60In1900IsTheFictitiousLeapDay() =>
        Assert.IsFalse(DateSerial.TryToDateTime(60.5, DateSystem.Date1900, out _));

    [TestMethod]
    [DataRow(-1.0, DateSystem.Date1900)]
    [DataRow(-0.5, DateSystem.Date1904)]
    [DataRow(2958466.0, DateSystem.Date1900)]
    [DataRow(2957004.0, DateSystem.Date1904)]
    [DataRow(double.NaN, DateSystem.Date1900)]
    public void RejectsSerialsOutsideTheRange(double serial, DateSystem system) =>
        Assert.IsFalse(DateSerial.TryToDateTime(serial, system, out _));

    [TestMethod]
    [DataRow(DateSystem.Date1900)]
    [DataRow(DateSystem.Date1904)]
    public void ReadsTheTimeOfDay(DateSystem system)
    {
        double serial = system == DateSystem.Date1900 ? 40729.75 : 39267.75;
        Assert.IsTrue(DateSerial.TryToDateTime(serial, system, out DateTime value));
        Assert.AreEqual(new DateTime(2011, 7, 5, 18, 0, 0), value);
    }

    [TestMethod]
    [DataRow(0.0, 0, 0, 0)]
    [DataRow(0.5, 12, 0, 0)]
    [DataRow(0.5625, 13, 30, 0)]
    public void ReadsASerialBelowOneAsATimeOnly(double serial, int hour, int minute, int second)
    {
        foreach (DateSystem system in new[] { DateSystem.Date1900, DateSystem.Date1904 })
        {
            Assert.IsTrue(DateSerial.TryToDateTime(serial, system, out DateTime value));
            Assert.IsTrue(DateSerial.IsTimeOnly(value));
            Assert.AreEqual(new TimeSpan(hour, minute, second), value.TimeOfDay);
        }
    }

    [TestMethod]
    public void RoundsToTheMillisecondAndCarriesIntoTheNextDay()
    {
        Assert.IsTrue(DateSerial.TryToDateTime(40729.999999999, DateSystem.Date1900, out DateTime value));
        Assert.AreEqual(new DateTime(2011, 7, 6), value);
    }

    [TestMethod]
    [DataRow(1900, 1, 1, 1.0)]
    [DataRow(1900, 2, 28, 59.0)]
    [DataRow(1900, 3, 1, 61.0)]
    [DataRow(2011, 7, 5, 40729.0)]
    [DataRow(9999, 12, 31, 2958465.0)]
    public void Writes1900Serials(int year, int month, int day, double serial)
    {
        Assert.IsTrue(DateSerial.TryFromDateTime(new DateTime(year, month, day), DateSystem.Date1900, out double value));
        Assert.AreEqual(serial, value);
    }

    [TestMethod]
    [DataRow(1904, 1, 2, 1.0)]
    [DataRow(2011, 7, 5, 39267.0)]
    public void Writes1904Serials(int year, int month, int day, double serial)
    {
        Assert.IsTrue(DateSerial.TryFromDateTime(new DateTime(year, month, day), DateSystem.Date1904, out double value));
        Assert.AreEqual(serial, value);
    }

    [TestMethod]
    public void TheDateSystemsDifferBy1462Days()
    {
        var date = new DateTime(2024, 1, 15, 13, 30, 0);
        Assert.IsTrue(DateSerial.TryFromDateTime(date, DateSystem.Date1900, out double serial1900));
        Assert.IsTrue(DateSerial.TryFromDateTime(date, DateSystem.Date1904, out double serial1904));
        Assert.AreEqual(1462.0, serial1900 - serial1904, 1e-9);
    }

    [TestMethod]
    [DataRow(1899, 12, 31, DateSystem.Date1900)]
    [DataRow(1903, 12, 31, DateSystem.Date1904)]
    public void RejectsDatesBeforeTheFirstDay(int year, int month, int day, DateSystem system) =>
        Assert.IsFalse(DateSerial.TryFromDateTime(new DateTime(year, month, day), system, out _));

    [TestMethod]
    public void WritesATimeOnlyAsAFractionOfADay()
    {
        Assert.IsTrue(DateSerial.TryFromDateTime(DateSerial.TimeOnlyDay.AddHours(18), DateSystem.Date1900, out double serial));
        Assert.AreEqual(0.75, serial);
    }

    [TestMethod]
    [DataRow(DateSystem.Date1900)]
    [DataRow(DateSystem.Date1904)]
    public void SerialsRoundTrip(DateSystem system)
    {
        foreach (double serial in new[] { 0.25, 1.0, 59.5, 61.0, 1462.0, 45306.5625, 45306.0, 2_000_000.125 })
        {
            Assert.IsTrue(DateSerial.TryToDateTime(serial, system, out DateTime value), serial.ToString(System.Globalization.CultureInfo.InvariantCulture));
            Assert.IsTrue(DateSerial.TryFromDateTime(value, system, out double back));
            Assert.AreEqual(serial, back, 1e-9);
        }
    }
}
