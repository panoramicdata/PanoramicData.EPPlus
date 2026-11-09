using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public class SubtotalTests : FormulaParserTestBase
{
	private ExcelWorksheet _worksheet;
	private ExcelPackage _package;

	[TestInitialize]
	public void Setup()
	{
		_package = new ExcelPackage(new MemoryStream());
		_worksheet = _package.Workbook.Worksheets.Add("Test");
		_parser = _worksheet.Workbook.FormulaParser;
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();

	private object CalculateOverNestedSubtotal(int function)
	{
		_worksheet.Cells["A1"].Formula = $"SUBTOTAL({function}, A2:A3)";
		_worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
		_worksheet.Cells["A3"].Value = 2d;
		_worksheet.Cells["A5"].Value = 2d;
		_worksheet.Calculate();
		return _worksheet.Cells["A1"].Value;
	}

	private object CalculateOverFourCells(string outerFormula, string innerFormula)
	{
		_worksheet.Cells["A1"].Formula = outerFormula;
		_worksheet.Cells["A2"].Formula = innerFormula;
		_worksheet.Cells["A3"].Value = 5d;
		_worksheet.Cells["A4"].Value = 4d;
		_worksheet.Cells["A5"].Value = 2d;
		_worksheet.Cells["A6"].Value = 4d;
		_worksheet.Calculate();
		return _worksheet.Cells["A1"].Value;
	}
	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_Avg() => Assert.AreEqual(2d, CalculateOverNestedSubtotal(1));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_Count() => Assert.AreEqual(1d, CalculateOverNestedSubtotal(2));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_CountA() => Assert.AreEqual(1d, CalculateOverNestedSubtotal(3));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_Max() => Assert.AreEqual(2d, CalculateOverNestedSubtotal(4));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_Min() => Assert.AreEqual(2d, CalculateOverNestedSubtotal(5));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_Product() => Assert.AreEqual(2d, CalculateOverNestedSubtotal(6));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_Stdev()
	{
		var result = CalculateOverFourCells("SUBTOTAL(7, A2:A4)", "SUBTOTAL(7, A5:A6)");
		result = Math.Round((double)result, 9);
		Assert.AreEqual(0.707106781d, result);
	}

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_StdevP() => Assert.AreEqual(0.5d, CalculateOverFourCells("SUBTOTAL(8, A2:A4)", "SUBTOTAL(8, A5:A6)"));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_Sum() => Assert.AreEqual(2d, CalculateOverNestedSubtotal(9));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_Var() => Assert.AreEqual(9d, CalculateOverFourCells("SUBTOTAL(9, A2:A4)", "SUBTOTAL(8, A5:A6)"));

	[TestMethod]
	public void SubtotalShouldNotIncludeSubtotalChildren_VarP() => Assert.AreEqual(0.5d, CalculateOverFourCells("SUBTOTAL(10, A2:A4)", "SUBTOTAL(8, A5:A6)"));
}
