using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class SubtotalTests
{
	[TestMethod]
	public void ShouldCalculateAverageWhenCalcTypeIs101()
	{
		var result = ExecuteOnSample(101);
		Assert.AreEqual(30d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateCountWhenCalcTypeIs102()
	{
		var result = ExecuteOnSample(102);
		Assert.AreEqual(5d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateCountAWhenCalcTypeIs103()
	{
		var result = ExecuteOnSample(103);
		Assert.AreEqual(5d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateMaxWhenCalcTypeIs104()
	{
		var result = ExecuteOnSample(104);
		Assert.AreEqual(50d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateMinWhenCalcTypeIs105()
	{
		var result = ExecuteOnSample(105);
		Assert.AreEqual(10d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateProductWhenCalcTypeIs106()
	{
		var result = ExecuteOnSample(106);
		Assert.AreEqual(12000000d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateStdevWhenCalcTypeIs107()
	{
		var result = ExecuteOnSample(107);
		var resultRounded = System.Math.Round((double)result.Result, 5);
		Assert.AreEqual(15.81139d, resultRounded);
	}

	[TestMethod]
	public void ShouldCalculateStdevPWhenCalcTypeIs108()
	{
		var result = ExecuteOnSample(108);
		var resultRounded = System.Math.Round((double)result.Result, 8);
		Assert.AreEqual(14.14213562, resultRounded);
	}

	[TestMethod]
	public void ShouldCalculateSumWhenCalcTypeIs109()
	{
		var result = ExecuteOnSample(109);
		Assert.AreEqual(150d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateVarWhenCalcTypeIs110()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(110, 10, 20, 30, 40, 50, 51);
		args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
		var result = func.Execute(args, _context);
		Assert.AreEqual(250d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateVarPWhenCalcTypeIs111()
	{
		var result = ExecuteOnSample(111);
		Assert.AreEqual(200d, result.Result);
	}
}
