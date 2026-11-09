using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class DateTimeFunctionsTests
{
	[TestMethod]
	public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs1()
	{
		var func = new Weekday();
		var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 1), _parsingContext);
		Assert.AreEqual(1, result.Result);
	}

	[TestMethod]
	public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs2()
	{
		var func = new Weekday();
		var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 2), _parsingContext);
		Assert.AreEqual(7, result.Result);
	}

	[TestMethod]
	public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs3()
	{
		var func = new Weekday();
		var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 3), _parsingContext);
		Assert.AreEqual(6, result.Result);
	}

	[TestMethod]
	public void WeekNumShouldReturnCorrectResult()
	{
		var func = new Weeknum();
		var dt1 = new DateTime(2012, 12, 31).ToOADate();
		var dt2 = new DateTime(2012, 1, 1).ToOADate();
		var dt3 = new DateTime(2013, 1, 20).ToOADate();

		var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1), _parsingContext);
		var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2), _parsingContext);
		var r3 = func.Execute(FunctionsHelper.CreateArgs(dt3, 2), _parsingContext);

		Assert.AreEqual(53, r1.Result, "r1.Result was not 53, but " + r1.Result.ToString());
		Assert.AreEqual(1, r2.Result, "r2.Result was not 1, but " + r2.Result.ToString());
		Assert.AreEqual(3, r3.Result, "r3.Result was not 3, but " + r3.Result.ToString());
	}

	[TestMethod]
	public void EdateShouldReturnCorrectResult()
	{
		var func = new Edate();

		var dt1arg = new DateTime(2012, 1, 31).ToOADate();
		var dt2arg = new DateTime(2013, 1, 1).ToOADate();
		var dt3arg = new DateTime(2013, 2, 28).ToOADate();

		var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1arg, 1), _parsingContext);
		var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2arg, -1), _parsingContext);
		var r3 = func.Execute(FunctionsHelper.CreateArgs(dt3arg, 2), _parsingContext);

		var dt1 = DateTime.FromOADate((double)r1.Result);
		var dt2 = DateTime.FromOADate((double)r2.Result);
		var dt3 = DateTime.FromOADate((double)r3.Result);

		var exp1 = new DateTime(2012, 2, 29);
		var exp2 = new DateTime(2012, 12, 1);
		var exp3 = new DateTime(2013, 4, 28);

		Assert.AreEqual(exp1, dt1, "dt1 was not " + exp1.ToString("yyyy-MM-dd") + ", but " + dt1.ToString("yyyy-MM-dd"));
		Assert.AreEqual(exp2, dt2, "dt1 was not " + exp2.ToString("yyyy-MM-dd") + ", but " + dt2.ToString("yyyy-MM-dd"));
		Assert.AreEqual(exp3, dt3, "dt1 was not " + exp3.ToString("yyyy-MM-dd") + ", but " + dt3.ToString("yyyy-MM-dd"));
	}

	[TestMethod]
	public void Days360ShouldReturnCorrectResultWithNoMethodSpecified2()
	{
		var func = new Days360();

		var dt1arg = new DateTime(2013, 1, 1).ToOADate();
		var dt2arg = new DateTime(2013, 3, 31).ToOADate();

		var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg), _parsingContext);

		Assert.AreEqual(90, result.Result);
	}

	[TestMethod]
	public void Days360ShouldReturnCorrectResultWithEuroMethodSpecified()
	{
		var func = new Days360();

		var dt1arg = new DateTime(2013, 1, 1).ToOADate();
		var dt2arg = new DateTime(2013, 3, 31).ToOADate();

		var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, true), _parsingContext);

		Assert.AreEqual(89, result.Result);
	}

	[TestMethod]
	public void Days360ShouldHandleFebWithEuroMethodSpecified()
	{
		var func = new Days360();

		var dt1arg = new DateTime(2012, 2, 29).ToOADate();
		var dt2arg = new DateTime(2013, 2, 28).ToOADate();

		var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, true), _parsingContext);

		Assert.AreEqual(359, result.Result);
	}

	[TestMethod]
	public void Days360ShouldHandleFebWithUsMethodSpecified()
	{
		var func = new Days360();

		var dt1arg = new DateTime(2012, 2, 29).ToOADate();
		var dt2arg = new DateTime(2013, 2, 28).ToOADate();

		var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, false), _parsingContext);

		Assert.AreEqual(358, result.Result);
	}

	[TestMethod]
	public void Days360ShouldHandleFebWithUsMethodSpecified2()
	{
		var func = new Days360();

		var dt1arg = new DateTime(2013, 2, 28).ToOADate();
		var dt2arg = new DateTime(2013, 3, 31).ToOADate();

		var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, false), _parsingContext);

		Assert.AreEqual(30, result.Result);
	}
}
