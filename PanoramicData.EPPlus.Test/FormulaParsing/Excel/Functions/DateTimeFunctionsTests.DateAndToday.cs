using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class DateTimeFunctionsTests
{
	[TestMethod]
	public void DateFunctionShouldReturnADate()
	{
		var func = new Date();
		var args = FunctionsHelper.CreateArgs(2012, 4, 3);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(DataType.Date, result.DataType);
	}

	[TestMethod]
	public void DateFunctionShouldReturnACorrectDate()
	{
		var expectedDate = new DateTime(2012, 4, 3);
		var func = new Date();
		var args = FunctionsHelper.CreateArgs(2012, 4, 3);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedDate.ToOADate(), result.Result);
	}

	[TestMethod]
	public void DateFunctionShouldMonthFromPrevYearIfMonthIsNegative()
	{
		var expectedDate = new DateTime(2011, 11, 3);
		var func = new Date();
		var args = FunctionsHelper.CreateArgs(2012, -1, 3);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedDate.ToOADate(), result.Result);
	}

	[TestMethod]
	public void NowFunctionShouldReturnNow()
	{
		var startTime = DateTime.Now;
		Thread.Sleep(1);
		var func = new Now();
		var args = Array.Empty<FunctionArgument>();
		var result = func.Execute(args, _parsingContext);
		Thread.Sleep(1);
		var endTime = DateTime.Now;
		var resultDate = DateTime.FromOADate((double)result.Result);
		Assert.IsTrue(resultDate > startTime && resultDate < endTime);
	}

	[TestMethod]
	public void TodayFunctionShouldReturnTodaysDate()
	{
		var func = new Today();
		var args = Array.Empty<FunctionArgument>();
		var result = func.Execute(args, _parsingContext);
		var resultDate = DateTime.FromOADate((double)result.Result);
		Assert.AreEqual(DateTime.Now.Date, resultDate);
	}

	[TestMethod]
	public void DayShouldReturnDayInMonth()
	{
		var date = new DateTime(2012, 3, 12);
		var func = new Day();
		var args = FunctionsHelper.CreateArgs(date.ToOADate());
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(12, result.Result);
	}

	[TestMethod]
	public void DayShouldReturnMonthOfYearWithStringParam()
	{
		var func = new Day();
		var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
		Assert.AreEqual(12, result.Result);
	}

	[TestMethod]
	public void MonthShouldReturnMonthOfYear()
	{
		var date = new DateTime(2012, 3, 12);
		var func = new Month();
		var result = func.Execute(FunctionsHelper.CreateArgs(date.ToOADate()), _parsingContext);
		Assert.AreEqual(3, result.Result);
	}

	[TestMethod]
	public void MonthShouldReturnMonthOfYearWithStringParam()
	{
		var func = new Month();
		var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
		Assert.AreEqual(3, result.Result);
	}

	[TestMethod]
	public void YearShouldReturnCorrectYear()
	{
		var date = new DateTime(2012, 3, 12);
		var func = new Year();
		var result = func.Execute(FunctionsHelper.CreateArgs(date.ToOADate()), _parsingContext);
		Assert.AreEqual(2012, result.Result);
	}

	[TestMethod]
	public void YearShouldReturnCorrectYearWithStringParam()
	{
		var func = new Year();
		var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
		Assert.AreEqual(2012, result.Result);
	}
}
