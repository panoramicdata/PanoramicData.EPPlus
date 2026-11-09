using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Exceptions;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class MathFunctionsTests
{
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
		AssertIndexSixOfFourThrows(new Large());
	}

	[TestMethod]
	public void SmallShouldReturnTheSmallestNumberIf1() => Assert.AreEqual(1d, ExecuteFunction(new Small(), FunctionsHelper.CreateArgs(1, 2, 3), 1));

	[TestMethod]
	public void SmallShouldReturnTheSecondSmallestNumberIf2() => Assert.AreEqual(2d, ExecuteFunction(new Small(), FunctionsHelper.CreateArgs(4, 1, 2, 3), 2));

	[TestMethod]
	public void SmallShouldThrowIfIndexOutOfBounds()
	{
		AssertIndexSixOfFourThrows(new Small());
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

	private void AssertIndexSixOfFourThrows(ExcelFunction func)
	{
		var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
		Assert.ThrowsExactly<ExcelErrorValueException>(() => func.Execute(args, _parsingContext));
	}
}
