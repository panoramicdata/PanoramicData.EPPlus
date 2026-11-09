using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public class MathFunctionsTests
{
	private ParsingContext _parsingContext;

	[TestInitialize]
	public void Initialize()
	{
		_parsingContext = ParsingContext.Create();
		_parsingContext.Scopes.NewScope(RangeAddress.Empty);
	}

	private object ExecuteFunction(ExcelFunction func, params object[] args)
		=> func.Execute(FunctionsHelper.CreateArgs(args), _parsingContext).Result;

	private double ExecuteFunctionRounded(ExcelFunction func, int digits, params object[] args)
		=> System.Math.Round((double)ExecuteFunction(func, args), digits);
	[TestMethod]
	public void PiShouldReturnPIConstant() => Assert.AreEqual(System.Math.Round(System.Math.PI, 14), ExecuteFunction(new Pi(), 0));

	[TestMethod]
	public void AbsShouldReturnCorrectResult() => Assert.AreEqual(3d, ExecuteFunction(new Abs(), -3d));

	[TestMethod]
	public void AsinShouldReturnCorrectResult() => Assert.AreEqual(1.5708, ExecuteFunctionRounded(new Asin(), 4, 1d));

	[TestMethod]
	public void AsinhShouldReturnCorrectResult() => Assert.AreEqual(0.0998, ExecuteFunctionRounded(new Asinh(), 4, 0.1d));

	[TestMethod]
	public void CeilingShouldRoundUpAccordingToParamsSignificanceLowerThan0() => Assert.AreEqual(22.36d, ExecuteFunction(new Ceiling(), 22.35d, 0.01));

	[TestMethod]
	public void CeilingBugTest1() => Assert.AreEqual(100d, ExecuteFunction(new Ceiling(), 100d, 100d));
	[TestMethod]
	public void CeilingBugTest2() => Assert.AreEqual(12000d, ExecuteFunction(new Ceiling(), 12000d, 1000d));

	[TestMethod]
	public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsMinus0point1()
	{
		var expectedValue = -22.4d;
		var func = new Ceiling();
		var args = FunctionsHelper.CreateArgs(-22.35d, -0.1);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedValue, System.Math.Round((double)result.Result, 2));
	}

	[TestMethod]
	public void CeilingShouldRoundUpAccordingToParamsSignificanceIs1() => Assert.AreEqual(23d, ExecuteFunction(new Ceiling(), 22.35d, 1));

	[TestMethod]
	public void CeilingShouldRoundUpAccordingToParamsSignificanceIs10() => Assert.AreEqual(30d, ExecuteFunction(new Ceiling(), 22.35d, 10));

	[TestMethod]
	public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsNegative() => Assert.AreEqual(-30d, ExecuteFunction(new Ceiling(), -22.35d, -10));

	[TestMethod]
	public void CeilingShouldThrowExceptionIfNumberIsPositiveAndSignificanceIsNegative()
	{
		var func = new Ceiling();
		var args = FunctionsHelper.CreateArgs(22.35d, -1);
		Assert.ThrowsExactly<InvalidOperationException>(() => func.Execute(args, _parsingContext));
	}

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

	[TestMethod]
	public void MaxShouldCalculateCorrectResult() => Assert.AreEqual(5d, ExecuteFunction(new Max(), 4, 2, 5, 2));

	[TestMethod]
	public void MaxShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
	{
		var func = new Max
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
		args.ElementAt(2).SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(4d, result.Result);
	}

	[TestMethod]
	public void MaxaShouldCalculateCorrectResult() => Assert.AreEqual(1d, ExecuteFunction(new Maxa(), -1, 0, 1));

	[TestMethod]
	public void MaxaShouldCalculateCorrectResultUsingBool() => Assert.AreEqual(1d, ExecuteFunction(new Maxa(), -1, 0, true));

	[TestMethod]
	public void MaxaShouldCalculateCorrectResultUsingString() => Assert.AreEqual(0d, ExecuteFunction(new Maxa(), -1, "test"));

	[TestMethod]
	public void MinShouldCalculateCorrectResult() => Assert.AreEqual(2d, ExecuteFunction(new Min(), 4, 2, 5, 2));

	[TestMethod]
	public void MinShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
	{
		var func = new Min
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(4, 2, 5, 3);
		args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(3d, result.Result);
	}

	[TestMethod]
	public void AverageShouldCalculateCorrectResult()
	{
		var expectedResult = (4d + 2d + 5d + 2d) / 4d;
		var func = new Average();
		var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedResult, result.Result);
	}

	[TestMethod]
	public void AverageShouldCalculateCorrectResultWithEnumerableAndBoolMembers()
	{
		var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
		var func = new Average();
		var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedResult, result.Result);
	}

	[TestMethod]
	public void AverageShouldIgnoreHiddenFieldsIfIgnoreHiddenValuesIsTrue()
	{
		var expectedResult = (4d + 2d + 2d + 1d) / 4d;
		var func = new Average
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
		args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedResult, result.Result);
	}

	[TestMethod]
	public void AverageShouldThrowDivByZeroExcelErrorValueIfEmptyArgs()
	{
		var errorType = eErrorType.Value;

		var func = new Average();
		var args = Array.Empty<FunctionArgument>();
		try
		{
			func.Execute(args, _parsingContext);
		}
		catch (ExcelErrorValueException e)
		{
			errorType = e.ErrorValue.Type;
		}

		Assert.AreEqual(eErrorType.Div0, errorType);
	}

	[TestMethod]
	public void AverageAShouldCalculateCorrectResult()
	{
		var expectedResult = (4d + 2d + 5d + 2d) / 4d;
		var func = new AverageA();
		var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedResult, result.Result);
	}

	[TestMethod]
	public void AverageAShouldIncludeTrueAs1()
	{
		var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
		var func = new AverageA();
		var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, true);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedResult, result.Result);
	}

	[TestMethod]
	public void AverageAShouldThrowValueExceptionIfNonNumericTextIsSupplied()
	{
		var func = new AverageA();
		var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, "ABC");
		Assert.ThrowsExactly<ExcelErrorValueException>(() => func.Execute(args, _parsingContext));
	}

	[TestMethod]
	public void AverageAShouldCountValueAs0IfNonNumericTextIsSuppliedInArray() => Assert.AreEqual(1.5d, ExecuteFunction(new AverageA(), FunctionsHelper.CreateArgs(1d, 2d, 3d, "ABC")));

	[TestMethod]
	public void AverageAShouldCountNumericStringWithValue() => Assert.AreEqual(5d, ExecuteFunction(new AverageA(), 4d, 2d, "9"));

	[TestMethod]
	public void RoundShouldReturnCorrectResult() => Assert.AreEqual(2.343d, ExecuteFunction(new Round(), 2.3433, 3));

	[TestMethod]
	public void RoundShouldReturnCorrectResultWhenNbrOfDecimalsIsNegative() => Assert.AreEqual(9000d, ExecuteFunction(new Round(), 9333, -3));

	[TestMethod]
	public void FloorShouldReturnCorrectResultWhenSignificanceIsBetween0And1() => Assert.AreEqual(26.7d, ExecuteFunction(new Floor(), 26.75d, 0.1));

	[TestMethod]
	public void FloorShouldReturnCorrectResultWhenSignificanceIs1() => Assert.AreEqual(26d, ExecuteFunction(new Floor(), 26.75d, 1));

	[TestMethod]
	public void FloorShouldReturnCorrectResultWhenSignificanceIsMinus1() => Assert.AreEqual(-26d, ExecuteFunction(new Floor(), -26.75d, -1));
	[TestMethod]
	public void FloorBugTest1() => Assert.AreEqual(100d, ExecuteFunction(new Floor(), 100d, 100d));
	[TestMethod]
	public void FloorBugTest2() => Assert.AreEqual(12000d, ExecuteFunction(new Floor(), 12000d, 1000d));

	[TestMethod]
	public void RandShouldReturnAValueBetween0and1()
	{
		var func = new Rand();
		var args = Array.Empty<FunctionArgument>();
		var result1 = func.Execute(args, _parsingContext);
		Assert.IsTrue((double)result1.Result is > 0 and < 1);
		var result2 = func.Execute(args, _parsingContext);
		Assert.AreNotEqual(result1.Result, result2.Result, "The two numbers were the same");
		Assert.IsTrue((double)result2.Result is > 0 and < 1);
	}

	[TestMethod]
	public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValues()
	{
		var func = new RandBetween();
		var args = FunctionsHelper.CreateArgs(1, 5);
		var result = func.Execute(args, _parsingContext);
		CollectionAssert.Contains(new List<double> { 1d, 2d, 3d, 4d, 5d }, result.Result);
	}

	[TestMethod]
	public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValuesWhenLowIsNegative()
	{
		var func = new RandBetween();
		var args = FunctionsHelper.CreateArgs(-5, 0);
		var result = func.Execute(args, _parsingContext);
		CollectionAssert.Contains(new List<double> { 0d, -1d, -2d, -3d, -4d, -5d }, result.Result);
	}

	[TestMethod]
	public void CountShouldReturnNumberOfNumericItems() => Assert.AreEqual(5d, ExecuteFunction(new Count(), 1d, 2m, 3, new DateTime(2012, 4, 1), "4"));

	[TestMethod]
	public void CountShouldIncludeNumericStringsAndDatesInArray() => Assert.AreEqual(5d, ExecuteFunction(new Count(), FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4")));


	[TestMethod]
	public void CountShouldIncludeEnumerableMembers() => Assert.AreEqual(3d, ExecuteFunction(new Count(), 1d, FunctionsHelper.CreateArgs(12, 13)));

	[TestMethod, Ignore]
	public void CountShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
	{
		var func = new Count
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
		args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(2d, result.Result);
	}

	[TestMethod]
	public void CountAShouldReturnNumberOfNonWhitespaceItems() => Assert.AreEqual(5d, ExecuteFunction(new CountA(), 1d, 2m, 3, new DateTime(2012, 4, 1), "4", null, string.Empty));

	[TestMethod]
	public void CountAShouldIncludeEnumerableMembers() => Assert.AreEqual(3d, ExecuteFunction(new CountA(), 1d, FunctionsHelper.CreateArgs(12, 13)));

	[TestMethod]
	public void CountAShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
	{
		var func = new CountA
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
		args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(2d, result.Result);
	}

	[TestMethod]
	public void ProductShouldMultiplyArguments() => Assert.AreEqual(16d, ExecuteFunction(new Product(), 2d, 2d, 4d));

	[TestMethod]
	public void ProductShouldHandleEnumerable() => Assert.AreEqual(32d, ExecuteFunction(new Product(), 2d, 2d, FunctionsHelper.CreateArgs(4d, 2d)));

	[TestMethod]
	public void ProductShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
	{
		var func = new Product
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
		args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(16d, result.Result);
	}

	[TestMethod]
	public void ProductShouldHandleFirstItemIsEnumerable() => Assert.AreEqual(32d, ExecuteFunction(new Product(), FunctionsHelper.CreateArgs(4d, 2d), 2d, 2d));

	[TestMethod]
	public void VarShouldReturnCorrectResult()
	{
		var func = new Var();
		var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
	}

	[TestMethod]
	public void VarShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
	{
		var func = new Var
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 9);
		args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
	}

	[TestMethod]
	public void VarPShouldReturnCorrectResult() => Assert.AreEqual(1.25d, ExecuteFunction(new VarP(), 1, 2, 3, 4));

	[TestMethod]
	public void VarPShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
	{
		var func = new VarP
		{
			IgnoreHiddenValues = true
		};
		var args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 9);
		args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(1.25d, result.Result);
	}

	[TestMethod]
	public void ModShouldReturnCorrectResult() => Assert.AreEqual(1d, ExecuteFunction(new Mod(), 5, 2));

	[TestMethod]
	public void CosShouldReturnCorrectResult() => Assert.AreEqual(-0.416146837d, ExecuteFunctionRounded(new Cos(), 9, 2));

	[TestMethod]
	public void CosHShouldReturnCorrectResult() => Assert.AreEqual(3.762195691, ExecuteFunctionRounded(new Cosh(), 9, 2));

	[TestMethod]
	public void AcosShouldReturnCorrectResult() => Assert.AreEqual(1.4706, ExecuteFunctionRounded(new Acos(), 4, 0.1));

	[TestMethod]
	public void ACosHShouldReturnCorrectResult() => Assert.AreEqual(1.317, ExecuteFunctionRounded(new Acosh(), 3, 2));

	[TestMethod]
	public void SinShouldReturnCorrectResult() => Assert.AreEqual(0.909297427, ExecuteFunctionRounded(new Sin(), 9, 2));

	[TestMethod]
	public void SinhShouldReturnCorrectResult() => Assert.AreEqual(3.626860408d, ExecuteFunctionRounded(new Sinh(), 9, 2));

	[TestMethod]
	public void TanShouldReturnCorrectResult() => Assert.AreEqual(-2.185039863d, ExecuteFunctionRounded(new Tan(), 9, 2));

	[TestMethod]
	public void TanhShouldReturnCorrectResult() => Assert.AreEqual(0.96402758d, ExecuteFunctionRounded(new Tanh(), 9, 2));

	[TestMethod]
	public void AtanShouldReturnCorrectResult() => Assert.AreEqual(1.471127674d, ExecuteFunctionRounded(new Atan(), 9, 10));

	[TestMethod]
	public void Atan2ShouldReturnCorrectResult() => Assert.AreEqual(1.107148718d, ExecuteFunctionRounded(new Atan2(), 9, 1, 2));

	[TestMethod]
	public void AtanhShouldReturnCorrectResult() => Assert.AreEqual(0.1003d, ExecuteFunctionRounded(new Atanh(), 4, 0.1));

	[TestMethod]
	public void LogShouldReturnCorrectResult() => Assert.AreEqual(0.301029996d, ExecuteFunctionRounded(new Log(), 9, 2));

	[TestMethod]
	public void LogShouldReturnCorrectResultWithBase() => Assert.AreEqual(1d, ExecuteFunction(new Log(), 2, 2));

	[TestMethod]
	public void Log10ShouldReturnCorrectResult() => Assert.AreEqual(0.301029996d, ExecuteFunctionRounded(new Log10(), 9, 2));

	[TestMethod]
	public void LnShouldReturnCorrectResult() => Assert.AreEqual(1.60944d, ExecuteFunctionRounded(new Ln(), 5, 5));

	[TestMethod]
	public void SqrtPiShouldReturnCorrectResult() => Assert.AreEqual(2.506628275d, ExecuteFunctionRounded(new SqrtPi(), 9, 2));

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

	[TestMethod]
	public void FactShouldRoundDownAndReturnCorrectResult() => Assert.AreEqual(120d, ExecuteFunction(new Fact(), 5.99));

	[TestMethod]
	public void FactShouldThrowWhenNegativeNumber()
	{
		var func = new Fact();
		var args = FunctionsHelper.CreateArgs(-1);
		Assert.ThrowsExactly<ExcelErrorValueException>(() => func.Execute(args, _parsingContext));
	}

	[TestMethod]
	public void QuotientShouldReturnCorrectResult() => Assert.AreEqual(2, ExecuteFunction(new Quotient(), 5, 2));

	[TestMethod]
	public void QuotientShouldThrowWhenDenomIs0()
	{
		var func = new Quotient();
		var args = FunctionsHelper.CreateArgs(1, 0);
		Assert.ThrowsExactly<ExcelErrorValueException>(() => func.Execute(args, _parsingContext));
	}

	[TestMethod]
	public void LargeShouldReturnTheLargestNumberIf1() => Assert.AreEqual(3d, ExecuteFunction(new Large(), FunctionsHelper.CreateArgs(1, 2, 3), 1));

	[TestMethod]
	public void LargeShouldReturnTheSecondLargestNumberIf2() => Assert.AreEqual(3d, ExecuteFunction(new Large(), FunctionsHelper.CreateArgs(4, 1, 2, 3), 2));

	[TestMethod]
	public void LargeShouldThrowIfIndexOutOfBounds()
	{
		var func = new Large();
		var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
		Assert.ThrowsExactly<ExcelErrorValueException>(() => func.Execute(args, _parsingContext));
	}

	[TestMethod]
	public void SmallShouldReturnTheSmallestNumberIf1() => Assert.AreEqual(1d, ExecuteFunction(new Small(), FunctionsHelper.CreateArgs(1, 2, 3), 1));

	[TestMethod]
	public void SmallShouldReturnTheSecondSmallestNumberIf2() => Assert.AreEqual(2d, ExecuteFunction(new Small(), FunctionsHelper.CreateArgs(4, 1, 2, 3), 2));

	[TestMethod]
	public void SmallShouldThrowIfIndexOutOfBounds()
	{
		var func = new Small();
		var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
		Assert.ThrowsExactly<ExcelErrorValueException>(() => func.Execute(args, _parsingContext));
	}

	[TestMethod]
	public void MedianShouldThrowIfNoArgs()
	{
		var func = new Median();
		var args = FunctionsHelper.Empty();
		Assert.ThrowsExactly<ExcelErrorValueException>(() => func.Execute(args, _parsingContext));
	}

	[TestMethod]
	public void MedianShouldCalculateCorrectlyWithOneMember() => Assert.AreEqual(1d, ExecuteFunction(new Median(), 1));

	[TestMethod]
	public void MedianShouldCalculateCorrectlyWithOddMembers() => Assert.AreEqual(3d, ExecuteFunction(new Median(), 3, 5, 1, 4, 2));

	[TestMethod]
	public void MedianShouldCalculateCorrectlyWithEvenMembers() => Assert.AreEqual(2.5d, ExecuteFunction(new Median(), 1, 2, 3, 4));

	[TestMethod]
	public void CountIfShouldHandleNegativeCriteria()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("test");
		sheet1.Cells["A1"].Value = -1;
		sheet1.Cells["A2"].Value = -2;
		sheet1.Cells["A3"].Formula = "CountIf(A1:A2,\"-1\")";
		sheet1.Calculate();
		Assert.AreEqual(1d, sheet1.Cells["A3"].Value);
	}
	[TestMethod]
	public void Rank()
	{
		using var p = new ExcelPackage();
		var w = p.Workbook.Worksheets.Add("testsheet");
		w.SetValue(1, 1, 1);
		w.SetValue(2, 1, 1);
		w.SetValue(3, 1, 2);
		w.SetValue(4, 1, 2);
		w.SetValue(5, 1, 4);
		w.SetValue(6, 1, 4);

		w.SetFormula(1, 2, "RANK(1,A1:A5)");
		w.SetFormula(1, 3, "RANK(1,A1:A5,1)");
		w.SetFormula(1, 4, "RANK.AVG(1,A1:A5)");
		w.SetFormula(1, 5, "RANK.AVG(1,A1:A5,1)");

		w.SetFormula(2, 2, "RANK.EQ(2,A1:A5)");
		w.SetFormula(2, 3, "RANK.EQ(2,A1:A5,1)");
		w.SetFormula(2, 4, "RANK.AVG(2,A1:A5,1)");
		w.SetFormula(2, 5, "RANK.AVG(2,A1:A5,0)");

		w.SetFormula(3, 2, "RANK(3,A1:A5)");
		w.SetFormula(3, 3, "RANK(3,A1:A5,1)");
		w.SetFormula(3, 4, "RANK.AVG(3,A1:A5,1)");
		w.SetFormula(3, 5, "RANK.AVG(3,A1:A5,0)");

		w.SetFormula(4, 2, "RANK.EQ(4,A1:A5)");
		w.SetFormula(4, 3, "RANK.EQ(4,A1:A5,1)");
		w.SetFormula(4, 4, "RANK.AVG(4,A1:A5,1)");
		w.SetFormula(4, 5, "RANK.AVG(4,A1:A5)");


		w.SetFormula(5, 4, "RANK.AVG(4,A1:A6,1)");
		w.SetFormula(5, 5, "RANK.AVG(4,A1:A6)");

		w.Calculate();

		Assert.AreEqual(4D, w.GetValue(1, 2));
		Assert.AreEqual(1D, w.GetValue(1, 3));
		Assert.AreEqual(4.5D, w.GetValue(1, 4));
		Assert.AreEqual(1.5D, w.GetValue(1, 5));

		Assert.AreEqual(2D, w.GetValue(2, 2));
		Assert.AreEqual(3D, w.GetValue(2, 3));
		Assert.AreEqual(3.5D, w.GetValue(2, 4));
		Assert.AreEqual(2.5D, w.GetValue(2, 5));

		Assert.IsInstanceOfType<ExcelErrorValue>(w.GetValue(3, 2));
		Assert.IsInstanceOfType<ExcelErrorValue>(w.GetValue(3, 3));
		Assert.IsInstanceOfType<ExcelErrorValue>(w.GetValue(3, 4));
		Assert.IsInstanceOfType<ExcelErrorValue>(w.GetValue(3, 5));

		Assert.AreEqual(1D, w.GetValue(4, 2));
		Assert.AreEqual(5D, w.GetValue(4, 3));
		Assert.AreEqual(5D, w.GetValue(4, 4));
		Assert.AreEqual(1D, w.GetValue(4, 5));

		Assert.AreEqual(5.5D, w.GetValue(5, 4));
		Assert.AreEqual(1.5D, w.GetValue(5, 5));
	}
}
