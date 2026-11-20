using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class DateTimeFunctionsTests
{
	[TestMethod]
	public void NetworkdaysShouldReturnNumberOfDays() => Assert.AreEqual(14, CalculateA1("NETWORKDAYS(DATE(2016,1,1), DATE(2016,1,20))"));

	[TestMethod]
	public void NetworkdaysShouldReturnNumberOfDaysWithHolidayRange()
	{
		using (MemoryStream ms = new())
		{
			// do something...
			using var package = new ExcelPackage();
			package.Load(ms);
		}

		using (var package = new ExcelPackage())
		{
			var ws = package.Workbook.Worksheets.Add("test");
			ws.Cells["A1"].Formula = "NETWORKDAYS(DATE(2016,1,1), DATE(2016,1,20),B1)";
			ws.Cells["B1"].Formula = "DATE(2016,1,15)";
			ws.Calculate();
			Assert.AreEqual(13, ws.Cells["A1"].Value);
		}
	}

	[TestMethod]
	public void NetworkdaysNegativeShouldReturnNumberOfDays() => Assert.AreEqual(10, CalculateA1("NETWORKDAYS(DATE(2016,1,1), DATE(2015,12,20))"));

	[TestMethod]
	public void NetworkdayIntlShouldUseWeekendArg() => Assert.AreEqual(17, CalculateA1("NETWORKDAYS.INTL(DATE(2016,1,1), DATE(2016,1,20), 11)"));

	[TestMethod]
	public void NetworkdayIntlShouldUseWeekendStringArg() => Assert.AreEqual(14, CalculateA1("NETWORKDAYS.INTL(DATE(2016,1,1), DATE(2016,1,20), \"0000011\")"));

	[TestMethod]
	public void NetworkdayIntlShouldReduceHoliday() => Assert.AreEqual(13, CalculateA1("NETWORKDAYS.INTL(DATE(2016,1,1), DATE(2016,1,20), \"0000011\", DATE(2016,1,4))"));

	[TestMethod]
	public void TimeAddition()
	{
		var a1 = CalculateA1("1 + (Time(10,0,0))");
		var result = Convert.ToDouble(a1);
		result = System.Math.Round(result, 2);
		Assert.AreEqual(1.42d, result);
	}
}
