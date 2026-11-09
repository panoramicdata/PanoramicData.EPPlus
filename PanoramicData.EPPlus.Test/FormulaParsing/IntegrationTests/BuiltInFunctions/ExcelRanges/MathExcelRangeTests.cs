using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges;

[TestClass]
public partial class MathExcelRangeTests
{
	private ExcelPackage _package;
	private ExcelWorksheet _worksheet;

	[TestInitialize]
	public void Initialize()
	{
		_package = new ExcelPackage();
		_worksheet = _package.Workbook.Worksheets.Add("Test");

		_worksheet.Cells["A1"].Value = 1;
		_worksheet.Cells["A2"].Value = 3;
		_worksheet.Cells["A3"].Value = 6;
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();

	private object Calculate(string address, string formula)
	{
		_worksheet.Cells[address].Formula = formula;
		_worksheet.Calculate();
		return _worksheet.Cells[address].Value;
	}
	[TestMethod]
	public void AbsShouldReturn3() => Assert.AreEqual(3d, Calculate("A4", "ABS(A2)"));

	[TestMethod]
	public void CountShouldReturn3() => Assert.AreEqual(3d, Calculate("A4", "COUNT(A1:A3)"));

	[TestMethod]
	public void CountShouldReturn2IfACellValueIsNull()
	{
		_worksheet.Cells["A2"].Value = null;
		var result = Calculate("A4", "COUNT(A1:A3)");
		Assert.AreEqual(2d, result);
	}

	[TestMethod]
	public void CountAShouldReturn3() => Assert.AreEqual(3d, Calculate("A4", "COUNTA(A1:A3)"));

	[TestMethod]
	public void CountIfShouldReturnCorrectResult() => Assert.AreEqual(2d, Calculate("A4", "COUNTIF(A1:A3, \">2\")"));

	[TestMethod]
	public void MaxShouldReturn6() => Assert.AreEqual(6d, Calculate("A4", "Max(A1:A3)"));

	[TestMethod]
	public void MinShouldReturn1() => Assert.AreEqual(1d, Calculate("A4", "Min(A1:A3)"));

	[TestMethod]
	public void AverageShouldReturn3Point333333() => Assert.AreEqual(3d + 1d / 3d, Calculate("A4", "Average(A1:A3)"));
}
