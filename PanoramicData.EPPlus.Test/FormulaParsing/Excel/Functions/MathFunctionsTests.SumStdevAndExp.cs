using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void SumShouldCalculate2Plus3AndReturn5() => Assert.AreEqual(5d, ExecuteFunction(new Sum(), 2, 3));

	[TestMethod]
	public void SumShouldCalculateEnumerableOf2Plus5Plus3AndReturn10() => Assert.AreEqual(10d, ExecuteFunction(new Sum(), FunctionsHelper.CreateArgs(2, 5), 3));

	[TestMethod]
	public void SumShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
	{
		var func = new Sum
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3, 4);
		args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(10d, result.Result);
	}

	[TestMethod]
	public void SumSqShouldCalculateArray() => Assert.AreEqual(20d, ExecuteFunction(new Sumsq(), 2, 4));

	[TestMethod]
	public void SumSqShouldIncludeTrueAsOne() => Assert.AreEqual(21d, ExecuteFunction(new Sumsq(), 2, 4, true));

	[TestMethod]
	public void SumSqShouldNoCountTrueTrueInArray() => Assert.AreEqual(20d, ExecuteFunction(new Sumsq(), FunctionsHelper.CreateArgs(2, 4, true)));

	[TestMethod]
	public void StdevShouldCalculateCorrectResult() => Assert.AreEqual(2d, ExecuteFunction(new Stdev(), 1, 3, 5));

	[TestMethod]
	public void StdevShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
	{
		var func = new Stdev
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(1, 3, 5, 6);
		args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(2d, result.Result);
	}

	[TestMethod]
	public void StdevPShouldCalculateCorrectResult()
	{
		var func = new StdevP();
		var args = FunctionsHelper.CreateArgs(2, 3, 4);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(0.8165d, System.Math.Round((double)result.Result, 5));
	}

	[TestMethod]
	public void StdevPShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
	{
		var func = new StdevP
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(2, 3, 4, 165);
		args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(0.8165d, System.Math.Round((double)result.Result, 5));
	}

	[TestMethod]
	public void ExpShouldCalculateCorrectResult()
	{
		var func = new Exp();
		var args = FunctionsHelper.CreateArgs(4);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(54.59815003d, System.Math.Round((double)result.Result, 8));
	}
}
