using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public partial class RefAndLookupTests : FormulaParserTestBase
{
	private ExcelDataProvider _excelDataProvider;
	const string WorksheetName = null;
	private ExcelPackage _package;
	private ExcelWorksheet _worksheet;

	[TestInitialize]
	public void Initialize()
	{
		_package = new ExcelPackage();
		_worksheet = _package.Workbook.Worksheets.Add("Test");
		_excelDataProvider = A.Fake<ExcelDataProvider>();
		A.CallTo(() => _excelDataProvider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(10, 1));
		A.CallTo(() => _excelDataProvider.GetWorkbookNameValues()).Returns(new ExcelNamedRangeCollection(_package.Workbook));
		_parser = new FormulaParser(_excelDataProvider);
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();

	private static readonly (string Address, object Value)[] OnesInColumnB = [("B1", 1d), ("B2", 1d), ("B3", 1d)];

	private static object CalculateOnNewSheet(string sheetName, string formulaCell, string formula, params (string Address, object Value)[] cells)
		=> SheetFormulaHelper.CalculateOnNewSheet(sheetName, formulaCell, formula, cells);

	[TestMethod]
	public void VLookupShouldReturnCorrespondingValue()
	{
		var lookupAddress = "A1:B2";
		var result = CalculateOnNewSheet("test", "A3", "VLOOKUP(2, " + lookupAddress + ", 2)",
			("A1", 1), ("B1", 1), ("A2", 2), ("B2", 5));
		Assert.AreEqual(5, result);
	}

	[TestMethod]
	public void VLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
	{
		var lookupAddress = "A1:B2";
		var result = CalculateOnNewSheet("test", "A3", "VLOOKUP(4, " + lookupAddress + ", 2, true)",
			("A1", 3), ("B1", 1), ("A2", 5), ("B2", 5));
		Assert.AreEqual(1, result);
	}

	[TestMethod]
	public void HLookupShouldReturnCorrespondingValue()
	{
		var lookupAddress = "A1:B2";
		_worksheet.Cells["A1"].Value = 1;
		_worksheet.Cells["B1"].Value = 2;
		_worksheet.Cells["A2"].Value = 2;
		_worksheet.Cells["B2"].Value = 5;
		_worksheet.Cells["A3"].Formula = "HLOOKUP(2, " + lookupAddress + ", 2)";
		_worksheet.Calculate();
		var result = _worksheet.Cells["A3"].Value;
		Assert.AreEqual(5, result);
	}

	[TestMethod]
	public void HLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
	{
		var lookupAddress = "A1:B2";
		var result = CalculateOnNewSheet("test", "E5", "HLOOKUP(4, " + lookupAddress + ", 2, true)",
			("A1", 3), ("B1", 5), ("A2", 1), ("B2", 2));
		Assert.AreEqual(1, result);
	}

	[TestMethod]
	public void LookupShouldReturnMatchingValue()
	{
		var lookupAddress = "A1:B2";
		var result = CalculateOnNewSheet("test", "E5", "LOOKUP(4, " + lookupAddress + ")",
			("A1", 3), ("B1", 5), ("A2", 4), ("B2", 1));
		Assert.AreEqual(1, result);
	}

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValue()
	{
		var lookupAddress = "A1:A2";
		A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
		A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName, 1, 2)).Returns(5);
		var result = _parser.Parse("MATCH(3, " + lookupAddress + ")");
		Assert.AreEqual(1, result);
	}
}