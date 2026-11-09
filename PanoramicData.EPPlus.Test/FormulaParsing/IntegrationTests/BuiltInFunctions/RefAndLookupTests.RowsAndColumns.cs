using FakeItEasy;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

public partial class RefAndLookupTests
{
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
}
