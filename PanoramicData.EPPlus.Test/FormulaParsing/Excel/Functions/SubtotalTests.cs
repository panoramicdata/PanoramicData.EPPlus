using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public partial class SubtotalTests
{
	private ParsingContext _context;

	[TestInitialize]
	public void Setup()
	{
		_context = ParsingContext.Create();
		_context.Scopes.NewScope(RangeAddress.Empty);
	}

	private CompileResult ExecuteOnSample(int functionNumber)
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(functionNumber, 10, 20, 30, 40, 50);
		return func.Execute(args, _context);
	}

	[TestMethod]
	public void ShouldThrowIfInvalidFuncNumber()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(139, 1);
		Assert.ThrowsExactly<ExcelErrorValueException>(() => func.Execute(args, _context));
	}

	[TestMethod]
	public void ShouldCalculateAverageWhenCalcTypeIs1()
	{
		var result = ExecuteOnSample(1);
		Assert.AreEqual(30d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateCountWhenCalcTypeIs2()
	{
		var result = ExecuteOnSample(2);
		Assert.AreEqual(5d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateCountAWhenCalcTypeIs3()
	{
		var result = ExecuteOnSample(3);
		Assert.AreEqual(5d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateMaxWhenCalcTypeIs4()
	{
		var result = ExecuteOnSample(4);
		Assert.AreEqual(50d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateMinWhenCalcTypeIs5()
	{
		var result = ExecuteOnSample(5);
		Assert.AreEqual(10d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateProductWhenCalcTypeIs6()
	{
		var result = ExecuteOnSample(6);
		Assert.AreEqual(12000000d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateStdevWhenCalcTypeIs7()
	{
		var result = ExecuteOnSample(7);
		var resultRounded = System.Math.Round((double)result.Result, 5);
		Assert.AreEqual(15.81139d, resultRounded);
	}

	[TestMethod]
	public void ShouldCalculateStdevPWhenCalcTypeIs8()
	{
		var result = ExecuteOnSample(8);
		var resultRounded = System.Math.Round((double)result.Result, 8);
		Assert.AreEqual(14.14213562, resultRounded);
	}

	[TestMethod]
	public void ShouldCalculateSumWhenCalcTypeIs9()
	{
		var result = ExecuteOnSample(9);
		Assert.AreEqual(150d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateVarWhenCalcTypeIs10()
	{
		var result = ExecuteOnSample(10);
		Assert.AreEqual(250d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateVarPWhenCalcTypeIs11()
	{
		var result = ExecuteOnSample(11);
		Assert.AreEqual(200d, result.Result);
	}
}
