using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class DateTimeFunctionsTests
{
	[TestMethod]
	public void YearFracShouldReturnCorrectResultWithUsBasis()
	{
		var result = ExecuteOnDates(new Yearfrac(), new DateTime(2013, 2, 28), new DateTime(2013, 3, 31));

		var roundedResult = System.Math.Round((double)result.Result, 4);

		Assert.IsTrue(System.Math.Abs(0.0861 - roundedResult) < double.Epsilon);
	}

	[TestMethod]
	public void YearFracShouldReturnCorrectResultWithEuroBasis()
	{
		var result = ExecuteOnDates(new Yearfrac(), new DateTime(2013, 2, 28), new DateTime(2013, 3, 31), 4);

		var roundedResult = System.Math.Round((double)result.Result, 4);

		Assert.IsTrue(System.Math.Abs(0.0889 - roundedResult) < double.Epsilon);
	}

	[TestMethod]
	public void YearFracActualActual()
	{
		var result = ExecuteOnDates(new Yearfrac(), new DateTime(2012, 2, 28), new DateTime(2013, 3, 31), 1);

		var roundedResult = System.Math.Round((double)result.Result, 4);

		Assert.IsTrue(System.Math.Abs(1.0862 - roundedResult) < double.Epsilon);
	}

	[TestMethod]
	public void IsoWeekShouldReturn1When1StJan()
	{
		var func = new IsoWeekNum();
		var arg = new DateTime(2013, 1, 1).ToOADate();

		var result = func.Execute(FunctionsHelper.CreateArgs(arg), _parsingContext);

		Assert.AreEqual(1, result.Result);
	}

	[TestMethod]
	public void EomonthShouldReturnCorrectResultWithPositiveArg()
	{
		var func = new Eomonth();
		var arg = new DateTime(2013, 2, 2).ToOADate();

		var result = func.Execute(FunctionsHelper.CreateArgs(arg, 3), _parsingContext);

		Assert.AreEqual(41425d, result.Result);
	}

	[TestMethod]
	public void EomonthShouldReturnCorrectResultWithNegativeArg()
	{
		var func = new Eomonth();
		var arg = new DateTime(2013, 2, 2).ToOADate();

		var result = func.Execute(FunctionsHelper.CreateArgs(arg, -3), _parsingContext);

		Assert.AreEqual(41243d, result.Result);
	}

	[TestMethod]
	public void WorkdayShouldReturnCorrectResultIfNoHolidayIsSupplied()
	{
		var expectedDate = new DateTime(2014, 1, 29).ToOADate();
		Assert.AreEqual(expectedDate, ExecuteWorkday(new DateTime(2014, 1, 1), 20));
	}

	[TestMethod]
	public void WorkdayShouldReturnCorrectResultWithNegativeArg()
	{
		var inputDate = new DateTime(2016, 6, 15).ToOADate();
		var expectedDate = new DateTime(2016, 5, 4).ToOADate();

		var func = new Workday();
		var args = FunctionsHelper.CreateArgs(inputDate, -30);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(DateTime.FromOADate(expectedDate), DateTime.FromOADate((double)result.Result));
	}

	[TestMethod]
	public void WorkdayShouldReturnCorrectResultWithFourDaysSupplied()
	{
		var expectedDate = new DateTime(2014, 1, 7).ToOADate();
		Assert.AreEqual(expectedDate, ExecuteWorkday(new DateTime(2014, 1, 1), 4));
	}

	[TestMethod]
	public void WorkdayWithNegativeArgShouldReturnCorrectWhenArrayOfHolidayDatesIsSupplied()
	{
		var inputDate = new DateTime(2016, 7, 27).ToOADate();
		var holidayDate1 = new DateTime(2016, 7, 11).ToOADate();
		var holidayDate2 = new DateTime(2016, 7, 8).ToOADate();
		var expectedDate = new DateTime(2016, 6, 13).ToOADate();

		var func = new Workday();
		var args = FunctionsHelper.CreateArgs(inputDate, -30, FunctionsHelper.CreateArgs(holidayDate1, holidayDate2));
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedDate, result.Result);
	}

	[TestMethod]
	public void WorkdayWithNegativeArgShouldReturnCorrectWhenRangeWithHolidayDatesIsSupplied()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("test");
		ws.Cells["A1"].Value = new DateTime(2016, 7, 27).ToOADate();
		ws.Cells["B1"].Value = new DateTime(2016, 7, 11).ToOADate();
		ws.Cells["B2"].Value = new DateTime(2016, 7, 8).ToOADate();
		ws.Cells["B3"].Formula = "WORKDAY(A1,-30, B1:B2)";
		ws.Calculate();

		var expectedDate = new DateTime(2016, 6, 13).ToOADate();
		var actualDate = ws.Cells["B3"].Value;
		Assert.AreEqual(expectedDate, actualDate);
	}
}
