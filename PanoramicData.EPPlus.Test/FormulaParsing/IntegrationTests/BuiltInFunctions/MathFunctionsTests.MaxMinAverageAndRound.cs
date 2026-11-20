using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void ExpShouldReturnAResult()
	{
		var result = _parser.Parse("exp(4)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void MaxShouldReturnAResult()
	{
		var result = _parser.Parse("Max(4, 5)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void MaxaShouldReturnAResult()
	{
		var result = _parser.Parse("Maxa(4, 5)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void MinShouldReturnAResult()
	{
		var result = _parser.Parse("min(4, 5)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void MinaShouldCalculateStringAs0()
	{
		using var pck = new ExcelPackage();
		var sheet = pck.Workbook.Worksheets.Add("test");
		sheet.Cells["A1"].Value = 1;
		sheet.Cells["B2"].Value = "a";
		sheet.Cells["A5"].Formula = "MINA(A1:B4)";
		sheet.Calculate();
		Assert.AreEqual(0d, sheet.Cells["A5"].Value);
	}

	[TestMethod]
	public void AverageShouldReturnAResult()
	{
		var result = _parser.Parse("Average(2, 2, 2)");
		Assert.AreEqual(2d, result);
	}

	[TestMethod]
	public void AverageShouldReturnDiv0IfEmptyCell()
	{
		using var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("test");
		ws.Cells["A2"].Formula = "AVERAGE(A1)";
		ws.Calculate();
		Assert.AreEqual("#DIV/0!", ws.Cells["A2"].Value.ToString());
	}

	[TestMethod]
	public void RoundShouldReturnAResult()
	{
		var result = _parser.Parse("Round(2.2, 0)");
		Assert.AreEqual(2d, result);
	}

	[TestMethod]
	public void RounddownShouldReturnAResult()
	{
		var result = _parser.Parse("Rounddown(2.99, 1)");
		Assert.AreEqual(2.9d, result);
	}

	[TestMethod]
	public void RoundupShouldReturnAResult()
	{
		var result = _parser.Parse("Roundup(2.99, 1)");
		Assert.AreEqual(3d, result);
	}

	[TestMethod]
	public void SqrtPiShouldReturnAResult()
	{
		var result = _parser.Parse("SqrtPi(2.2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void IntShouldReturnAResult()
	{
		var result = _parser.Parse("Int(2.9)");
		Assert.AreEqual(2, result);
	}
}
