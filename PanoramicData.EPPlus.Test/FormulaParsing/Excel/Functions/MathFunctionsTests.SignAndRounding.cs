using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void SignShouldReturnMinus1IfArgIsNegative()
	{
		var func = new Sign();
		var args = FunctionsHelper.CreateArgs(-2);
		var result = func.Execute(args, _parsingContext).Result;
		Assert.AreEqual(-1d, result);
	}

	[TestMethod]
	public void SignShouldReturn1IfArgIsPositive()
	{
		var func = new Sign();
		var args = FunctionsHelper.CreateArgs(2);
		var result = func.Execute(args, _parsingContext).Result;
		Assert.AreEqual(1d, result);
	}

	[TestMethod]
	public void RounddownShouldReturnCorrectResultWithPositiveNumber() => Assert.AreEqual(9.99, ExecuteFunction(new Rounddown(), 9.999, 2));

	[TestMethod]
	public void RounddownShouldHandleNegativeNumber() => Assert.AreEqual(-9.99, ExecuteFunction(new Rounddown(), -9.999, 2));

	[TestMethod]
	public void RounddownShouldHandleNegativeNumDigits() => Assert.AreEqual(900d, ExecuteFunction(new Rounddown(), 999.999, -2));

	[TestMethod]
	public void RounddownShouldReturn0IfNegativeNumDigitsIsTooLarge() => Assert.AreEqual(0d, ExecuteFunction(new Rounddown(), 999.999, -4));

	[TestMethod]
	public void RounddownShouldHandleZeroNumDigits() => Assert.AreEqual(999d, ExecuteFunction(new Rounddown(), 999.999, 0));

	[TestMethod]
	public void RoundupShouldReturnCorrectResultWithPositiveNumber() => Assert.AreEqual(9.992, ExecuteFunction(new Roundup(), 9.9911, 3));

	[TestMethod]
	public void RoundupShouldHandleNegativeNumDigits() => Assert.AreEqual(99200d, ExecuteFunction(new Roundup(), 99123, -2));

	[TestMethod]
	public void RoundupShouldHandleZeroNumDigits() => Assert.AreEqual(1000d, ExecuteFunction(new Roundup(), 999.999, 0));

	[TestMethod]
	public void TruncShouldReturnCorrectResult() => Assert.AreEqual(99d, ExecuteFunction(new Trunc(), 99.99));
}
