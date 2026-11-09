using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;
using PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public class RefAndLookupTests : FormulaParserTestBase
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
	{
		using var package = new ExcelPackage();
		var sheet = package.Workbook.Worksheets.Add(sheetName);
		foreach (var (address, value) in cells)
		{
			sheet.Cells[address].Value = value;
		}

		sheet.Cells[formulaCell].Formula = formula;
		sheet.Calculate();
		return sheet.Cells[formulaCell].Value;
	}

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

	[TestMethod]
	public void RowShouldReturnRowNumber()
	{
		A.CallTo(() => _excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Row()");
		var result = _parser.ParseAt("A4");
		Assert.AreEqual(4, result);
	}

	[TestMethod]
	public void RowSholdHandleReference() => Assert.AreEqual(4, CalculateOnNewSheet("test", "A1", "ROW(A4)"));

	[TestMethod]
	public void ColumnShouldReturnRowNumber()
	{
		A.CallTo(() => _excelDataProvider.GetRangeFormula("", 4, 2)).Returns("Column()");
		var result = _parser.ParseAt("B4");
		Assert.AreEqual(2, result);
	}

	[TestMethod]
	public void ColumnSholdHandleReference() => Assert.AreEqual(2, CalculateOnNewSheet("test", "A1", "COLUMN(B4)"));

	[TestMethod]
	public void RowsShouldReturnNbrOfRows()
	{
		A.CallTo(() => _excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Rows(A5:B7)");
		var result = _parser.ParseAt("A4");
		Assert.AreEqual(3, result);
	}

	[TestMethod]
	public void ColumnsShouldReturnNbrOfCols()
	{
		A.CallTo(() => _excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Columns(A5:B7)");
		var result = _parser.ParseAt("A4");
		Assert.AreEqual(2, result);
	}

	[TestMethod]
	public void ChooseShouldReturnCorrectResult()
	{
		var result = _parser.Parse("Choose(1, \"A\", \"B\")");
		Assert.AreEqual("A", result);
	}

	[TestMethod]
	public void AddressShouldReturnCorrectResult()
	{
		A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(12345);
		var result = _parser.Parse("Address(1, 1)");
		Assert.AreEqual("$A$1", result);
	}

	[TestMethod]
	public void IndirectShouldReturnARange()
	{
		using var package = new ExcelPackage(new MemoryStream());
		var s1 = package.Workbook.Worksheets.Add("Test");
		s1.Cells["A1:A2"].Value = 2;
		s1.Cells["A3"].Formula = "SUM(Indirect(\"A1:A2\"))";
		s1.Calculate();
		Assert.AreEqual(4d, s1.Cells["A3"].Value);

		s1.Cells["A4"].Formula = "SUM(Indirect(\"A1:A\" & \"2\"))";
		s1.Calculate();
		Assert.AreEqual(4d, s1.Cells["A4"].Value);
	}

	[TestMethod]
	public void OffsetShouldReturnASingleValue()
		=> Assert.AreEqual(1d, CalculateOnNewSheet("Test", "A5", "OFFSET(A1, 2, 1)", ("B3", 1d)));

	[TestMethod]
	public void OffsetShouldReturnARange()
		=> Assert.AreEqual(3d, CalculateOnNewSheet("Test", "A5", "SUM(OFFSET(A1:A3, 0, 1))", OnesInColumnB));

	[TestMethod]
	public void OffsetShouldReturnARange2()
		=> Assert.AreEqual(3d, CalculateOnNewSheet("Test", "A5", "COUNTA(OFFSET(Test!B1, 0, 0, Test!B2, Test!B3))",
			("B1", 10d), ("B2", 10d), ("B3", 10d)));

	[TestMethod]
	public void OffsetDirectReferenceToMultiRangeShouldSetValueError()
	{
		var result = CalculateOnNewSheet("Test", "A5", "OFFSET(A1:A3, 0, 1)", OnesInColumnB);
		Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
	}

	[TestMethod]
	public void OffsetShouldReturnARangeAccordingToWidth()
		=> Assert.AreEqual(2d, CalculateOnNewSheet("Test", "A5", "SUM(OFFSET(A1:A3, 0, 1, 2))", OnesInColumnB));

	[TestMethod]
	public void OffsetShouldReturnARangeAccordingToHeight()
		=> Assert.AreEqual(6d, CalculateOnNewSheet("Test", "A5", "SUM(OFFSET(A1:A3, 0, 1, 2, 2))",
			[.. OnesInColumnB, ("C1", 2d), ("C2", 2d), ("C3", 2d)]));

	[TestMethod]
	public void OffsetShouldCoverMultipleColumns()
		=> Assert.AreEqual(9d, CalculateOnNewSheet("Test", "A5", "SUM(OFFSET(A1:B3, 0, 2))",
			("C1", 1d), ("C2", 1d), ("C3", 1d), ("D1", 2d), ("D2", 2d), ("D3", 2d)));

	[TestMethod, Ignore]
	public void VLookupShouldHandleNames()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\Book3.xlsx"));
		var s1 = package.Workbook.Worksheets.First();
		_ = s1.Cells["X10"].Formula;
	}

	[TestMethod]
	public void LookupShouldReturnFromResultVector()
	{
		var lookupAddress = "A1:A5";
		var resultAddress = "B1:B5";
		var result = CalculateOnNewSheet("test", "E5", "LOOKUP(C1, " + lookupAddress + ", " + resultAddress + ")",
			("A1", 4.14), ("A2", 4.19), ("A3", 5.17), ("A4", 5.77), ("A5", 6.39),
			("B1", "red"), ("B2", "orange"), ("B3", "yellow"), ("B4", "green"), ("B5", "blue"),
			("C1", 4.14));
		Assert.AreEqual("red", result);
	}
}