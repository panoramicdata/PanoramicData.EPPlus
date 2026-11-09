using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

public partial class DateAndTimeFunctionsTests
{
	[TestMethod]
	public void Day360ShouldReturnCorrectResult()
	{
		var result = _parser.Parse("Days360(Date(2012, 4, 2), Date(2012, 5, 2))");
		Assert.AreEqual(30, result);
	}

	[TestMethod]
	public void YearfracShouldReturnAResult()
	{
		var result = _parser.Parse("Yearfrac(Date(2012, 4, 2), Date(2012, 5, 2))");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void IsoWeekNumShouldReturnAResult()
	{
		var result = _parser.Parse("IsoWeekNum(Date(2012, 4, 2))");
		Assert.IsInstanceOfType<int>(result);
	}

	[TestMethod]
	public void EomonthShouldReturnAResult()
	{
		var result = _parser.Parse("Eomonth(Date(2013, 2, 2), 3)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void WorkdayShouldReturnAResult()
	{
		var result = _parser.Parse("Workday(Date(2013, 2, 2), 3)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void DateNotEqualToStringShouldBeTrue()
	{
		var result = _parser.Parse("TODAY() <> \"\"");
		Assert.IsTrue((bool)result);
	}

	[TestMethod]
	public void Calculation5()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("Calc1");
		ws.Cells["A1"].Value = "John";
		ws.Cells["B1"].Value = "Doe";
		ws.Cells["C1"].Formula = "B1&\", \"&A1";
		ws.Calculate();
		Assert.AreEqual("Doe, John", ws.Cells["C1"].Value);
	}

	[TestMethod]
	public void HourWithExcelReference()
	{
		Assert.AreEqual(10, CalculateOnTenElevenTwelve("HOUR(A1)"));
	}

	[TestMethod]
	public void MinuteWithExcelReference()
	{
		Assert.AreEqual(11, CalculateOnTenElevenTwelve("MINUTE(A1)"));
	}

	[TestMethod]
	public void SecondWithExcelReference()
	{
		Assert.AreEqual(12, CalculateOnTenElevenTwelve("SECOND(A1)"));
	}

	private static object CalculateOnTenElevenTwelve(string formula)
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("Calc1");
		ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
		ws.Cells["B1"].Formula = formula;
		ws.Calculate();
		return ws.Cells["B1"].Value;
	}
}
