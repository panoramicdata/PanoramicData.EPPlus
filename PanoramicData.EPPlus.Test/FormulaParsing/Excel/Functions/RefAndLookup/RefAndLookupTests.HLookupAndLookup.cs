using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

public partial class RefAndLookupTests
{
	[TestMethod]
	public void HLookupShouldReturnResultFromMatchingRow()
	{
		var result = ExecuteOnFakeCells(new HLookup(), [2, "A1:B2", 2], false,
			(1, 1, 1), (1, 2, 1), (2, 1, 2), (2, 2, 5));
		Assert.AreEqual(5, result.Result);
	}

	[TestMethod]
	public void HLookupShouldReturnNaErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsFalse()
	{
		var result = ExecuteOnFakeCells(new HLookup(), [2, "A1:B2", 2, false], false,
			(1, 1, 3), (1, 2, 1), (2, 1, 2), (2, 2, 5));
		var expectedResult = ExcelErrorValue.Create(eErrorType.NA);
		Assert.AreEqual(expectedResult, result.Result);
	}

	[TestMethod]
	public void HLookupShouldReturnErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsTrue()
	{
		var result = ExecuteOnFakeCells(new HLookup(), [1, "A1:B2", 2, true], false,
			(1, 1, 2), (1, 2, 3), (2, 1, 3), (2, 2, 5));
		Assert.AreEqual(DataType.ExcelError, result.DataType);
	}

	[TestMethod]
	public void LookupShouldReturnResultFromMatchingRowArrayVertical()
	{
		var result = ExecuteOnFakeCells(new Lookup(), [4, "A1:B3", 2], true,
			(1, 1, 1), (1, 2, "A"), (2, 1, 3), (2, 2, "B"), (3, 1, 5), (3, 2, "C"));
		Assert.AreEqual("B", result.Result);
	}

	[TestMethod]
	public void LookupShouldReturnResultFromMatchingRowArrayHorizontal()
	{
		var result = ExecuteOnFakeCells(new Lookup(), [4, "A1:C2", 2], false,
			(1, 1, 1), (1, 2, 3), (1, 3, 5), (2, 1, "A"), (2, 2, "B"), (2, 3, "C"));
		Assert.AreEqual("B", result.Result);
	}

	[TestMethod]
	public void LookupShouldReturnResultFromMatchingSecondArrayHorizontal()
	{
		var result = CalculateLookupOnSheet("LOOKUP(4, A1:C1, A3:C3)",
			("A1", 1), ("B1", 3), ("C1", 5), ("A3", "A"), ("B3", "B"), ("C3", "C"));
		Assert.AreEqual("B", result);
	}

	[TestMethod]
	public void LookupShouldReturnResultFromMatchingSecondArrayHorizontalWithOffset()
	{
		var result = CalculateLookupOnSheet("LOOKUP(4, A1:C1, B3:D3)",
			("A1", 1), ("B1", 3), ("C1", 5), ("B3", "A"), ("C3", "B"), ("D3", "C"));
		Assert.AreEqual("B", result);
	}
}
