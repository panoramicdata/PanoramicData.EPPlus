using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class DateTimeFunctionsTests
{
	[TestMethod]
	public void TimeShouldReturnACorrectSerialNumber()
	{
		var expectedResult = GetTime(10, 11, 12);
		var func = new Time();
		var result = func.Execute(FunctionsHelper.CreateArgs(10, 11, 12), _parsingContext);
		Assert.AreEqual(expectedResult, result.Result);
	}

	[TestMethod]
	public void TimeShouldParseStringCorrectly()
	{
		var expectedResult = GetTime(10, 11, 12);
		var func = new Time();
		var result = func.Execute(FunctionsHelper.CreateArgs("10:11:12"), _parsingContext);
		Assert.AreEqual(expectedResult, result.Result);
	}

	[TestMethod]
	public void TimeShouldThrowExceptionIfSecondsIsOutOfRange()
	{
		var func = new Time();
		Assert.ThrowsExactly<ArgumentException>(() => func.Execute(FunctionsHelper.CreateArgs(10, 11, 60), _parsingContext));
	}

	[TestMethod]
	public void TimeShouldThrowExceptionIfMinuteIsOutOfRange()
	{
		var func = new Time();
		Assert.ThrowsExactly<ArgumentException>(() => func.Execute(FunctionsHelper.CreateArgs(10, 60, 12), _parsingContext));
	}

	[TestMethod]
	public void TimeShouldThrowExceptionIfHourIsOutOfRange()
	{
		var func = new Time();
		Assert.ThrowsExactly<ArgumentException>(() => func.Execute(FunctionsHelper.CreateArgs(24, 12, 12), _parsingContext));
	}

	[TestMethod]
	public void HourShouldReturnCorrectResult()
	{
		var func = new Hour();
		var result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 13, 14)), _parsingContext);
		Assert.AreEqual(9, result.Result);

		result = func.Execute(FunctionsHelper.CreateArgs(GetTime(23, 13, 14)), _parsingContext);
		Assert.AreEqual(23, result.Result);
	}

	[TestMethod]
	public void MinuteShouldReturnCorrectResult()
	{
		var func = new Minute();
		var result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 14, 14)), _parsingContext);
		Assert.AreEqual(14, result.Result);

		result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 55, 14)), _parsingContext);
		Assert.AreEqual(55, result.Result);
	}

	[TestMethod]
	public void SecondShouldReturnCorrectResult()
	{
		var func = new Second();
		var result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 14, 17)), _parsingContext);
		Assert.AreEqual(17, result.Result);
	}

	[TestMethod]
	public void SecondShouldReturnCorrectResultWithStringArgument()
	{
		var func = new Second();
		var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
		Assert.AreEqual(12, result.Result);
	}

	[TestMethod]
	public void MinuteShouldReturnCorrectResultWithStringArgument()
	{
		var func = new Minute();
		var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
		Assert.AreEqual(11, result.Result);
	}

	[TestMethod]
	public void HourShouldReturnCorrectResultWithStringArgument()
	{
		var func = new Hour();
		var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
		Assert.AreEqual(10, result.Result);
	}
}
