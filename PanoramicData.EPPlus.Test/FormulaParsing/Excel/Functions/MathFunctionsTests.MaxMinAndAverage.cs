using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Exceptions;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class MathFunctionsTests
{
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
}
