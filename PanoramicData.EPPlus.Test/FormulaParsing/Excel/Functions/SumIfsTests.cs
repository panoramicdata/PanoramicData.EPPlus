using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public class SumIfsTests
{
	private ExcelPackage _package;
	private ExcelWorksheet _sheet;
	[TestInitialize]
	public void Initialize()
	{
		_package = new ExcelPackage();
		var s1 = ThreeColumnSheet.Add(_package);
		_sheet = s1;
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();

	[TestMethod]
	public void ShouldCalculateTwoCriteriaRanges()
	{
		_sheet.Cells["A5"].Formula = "SUMIFS(A1:A4;B1:B5;\">5\";C1:C5;\">4\")";
		_sheet.Calculate();

		Assert.AreEqual(9d, _sheet.Cells["A5"].Value);
	}

	[TestMethod]
	public void ShouldIgnoreErrorInCriteriaRange()
	{
		_sheet.Cells["B3"].Value = ExcelErrorValue.Create(eErrorType.Div0);

		_sheet.Cells["A5"].Formula = "SUMIFS(A1:A4;B1:B5;\">5\";C1:C5;\">4\")";
		_sheet.Calculate();

		Assert.AreEqual(6d, _sheet.Cells["A5"].Value);
	}

	[TestMethod]
	public void ShouldHandleExcelRangesInCriteria()
	{
		_sheet.Cells["D1"].Value = 6;
		_sheet.Cells["A5"].Formula = "SUMIFS(A1:A4;B1:B5;\">5\";C1:C5;D1)";
		_sheet.Calculate();

		Assert.AreEqual(2d, _sheet.Cells["A5"].Value);
	}
}
