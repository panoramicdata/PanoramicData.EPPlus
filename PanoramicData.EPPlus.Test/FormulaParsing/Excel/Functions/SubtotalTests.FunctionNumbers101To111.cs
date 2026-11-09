using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class SubtotalTests
{
	[TestMethod]
	public void ShouldCalculateAverageWhenCalcTypeIs101()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(101, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		Assert.AreEqual(30d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateCountWhenCalcTypeIs102()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(102, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		Assert.AreEqual(5d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateCountAWhenCalcTypeIs103()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(103, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		Assert.AreEqual(5d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateMaxWhenCalcTypeIs104()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(104, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		Assert.AreEqual(50d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateMinWhenCalcTypeIs105()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(105, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		Assert.AreEqual(10d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateProductWhenCalcTypeIs106()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(106, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		Assert.AreEqual(12000000d, result.Result);
	}

	[TestMethod]
	public void ShouldCalculateStdevWhenCalcTypeIs107()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(107, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		var resultRounded = System.Math.Round((double)result.Result, 5);
		Assert.AreEqual(15.81139d, resultRounded);
	}

	[TestMethod]
	public void ShouldCalculateStdevPWhenCalcTypeIs108()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(108, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		var resultRounded = System.Math.Round((double)result.Result, 8);
		Assert.AreEqual(14.14213562, resultRounded);
	}

	[TestMethod]
	public void ShouldCalculateSumWhenCalcTypeIs109()
	{
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(109, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
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
		var func = new Subtotal();
		var args = FunctionsHelper.CreateArgs(111, 10, 20, 30, 40, 50);
		var result = func.Execute(args, _context);
		Assert.AreEqual(200d, result.Result);
	}
}
