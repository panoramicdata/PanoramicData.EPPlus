using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using AddressFunction = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Address;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

[TestClass]
public class RefAndLookupTests
{
	const string WorksheetName = null;

	private static LookupArguments CreateLookupArguments(params object[] args)
		=> new(FunctionsHelper.CreateArgs(args), ParsingContext.Create());

	private static CompileResult ExecuteOnFakeCells(ExcelFunction func, object[] args, bool fakeDimension, params (int Row, int Col, object Value)[] cells)
	{
		var parsingContext = ParsingContext.Create();
		parsingContext.Scopes.NewScope(RangeAddress.Empty);

		var provider = A.Fake<ExcelDataProvider>();
		foreach (var (row, col, value) in cells)
		{
			A.CallTo(() => provider.GetCellValue(WorksheetName, row, col)).Returns(value);
		}

		if (fakeDimension)
		{
			A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));
		}

		parsingContext.ExcelDataProvider = provider;
		return func.Execute(FunctionsHelper.CreateArgs(args), parsingContext);
	}

	private static ParsingContext CreateContextWithFakeProvider()
	{
		var parsingContext = ParsingContext.Create();
		parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
		return parsingContext;
	}

	private static ParsingContext CreateContextInScope(string address)
	{
		var parsingContext = ParsingContext.Create();
		var rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>());
		parsingContext.Scopes.NewScope(rangeAddressFactory.Create(address));
		return parsingContext;
	}

	private static ParsingContext CreateAddressContext()
	{
		var parsingContext = CreateContextWithFakeProvider();
		A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
		return parsingContext;
	}

	private static object CalculateLookupOnSheet(string formula, params (string Address, object Value)[] cells)
	{
		using var package = new ExcelPackage();
		var sheet = package.Workbook.Worksheets.Add("test");
		foreach (var (address, value) in cells)
		{
			sheet.Cells[address].Value = value;
		}

		sheet.Cells["D1"].Formula = formula;
		sheet.Calculate();
		return sheet.Cells["D1"].Value;
	}

	[TestMethod]
	public void LookupArgumentsShouldSetSearchedValue()
		=> Assert.AreEqual(1, CreateLookupArguments(1, "A:B", 2).SearchedValue);

	[TestMethod]
	public void LookupArgumentsShouldSetRangeAddress()
		=> Assert.AreEqual("A:B", CreateLookupArguments(1, "A:B", 2).RangeAddress);

	[TestMethod]
	public void LookupArgumentsShouldSetColIndex()
		=> Assert.AreEqual(2, CreateLookupArguments(1, "A:B", 2).LookupIndex);

	[TestMethod]
	public void LookupArgumentsShouldSetRangeLookupToTrueAsDefaultValue()
		=> Assert.IsTrue(CreateLookupArguments(1, "A:B", 2).RangeLookup);

	[TestMethod]
	public void LookupArgumentsShouldSetRangeLookupToTrueWhenTrueIsSupplied()
		=> Assert.IsTrue(CreateLookupArguments(1, "A:B", 2, true).RangeLookup);

	[TestMethod]
	public void VLookupShouldReturnResultFromMatchingRow()
	{
		var result = ExecuteOnFakeCells(new VLookup(), [2, "A1:B2", 2], true,
			(1, 1, 1), (1, 2, 1), (2, 1, 2), (2, 2, 5));
		Assert.AreEqual(5, result.Result);
	}

	[TestMethod]
	public void VLookupShouldReturnClosestValueBelowWhenRangeLookupIsTrue()
	{
		var result = ExecuteOnFakeCells(new VLookup(), [4, "A1:B2", 2, true], false,
			(1, 1, 3), (1, 2, 1), (2, 1, 5), (2, 2, 4));
		Assert.AreEqual(1, result.Result);
	}

	[TestMethod]
	public void VLookupShouldReturnClosestStringValueBelowWhenRangeLookupIsTrue()
	{
		var result = ExecuteOnFakeCells(new VLookup(), ["B", "A1:B2", 2, true], false,
			(1, 1, "A"), (1, 2, 1), (2, 1, "C"), (2, 2, 4));
		Assert.AreEqual(1, result.Result);
	}

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

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeExact()
	{
		var result = ExecuteOnFakeCells(new Match(), [3, "A1:C1", 0], false,
			(1, 1, 1), (1, 2, 3), (1, 3, 5));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValVertical_MatchTypeExact()
	{
		var result = ExecuteOnFakeCells(new Match(), [3, "A1:A3", 0], true,
			(1, 1, 1), (2, 1, 3), (3, 1, 5));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestBelow()
	{
		var result = ExecuteOnFakeCells(new Match(), [4, "A1:C1", 1], false,
			(1, 1, 1), (1, 2, 3), (1, 3, 5));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestAbove()
	{
		var result = ExecuteOnFakeCells(new Match(), [6, "A1:C1", -1], false,
			(1, 1, 10), (1, 2, 8), (1, 3, 5));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnFirstItemWhenExactMatch_MatchTypeClosestAbove()
	{
		var result = ExecuteOnFakeCells(new Match(), [10, "A1:C1", -1], false,
			(1, 1, 10), (1, 2, 8), (1, 3, 5));
		Assert.AreEqual(1, result.Result);
	}

	[TestMethod]
	public void MatchShouldHandleAddressOnOtherSheet()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
		sheet1.Cells["A1"].Formula = "Match(10, Sheet2!A1:Sheet2!A3, 0)";
		sheet2.Cells["A1"].Value = 9;
		sheet2.Cells["A2"].Value = 10;
		sheet2.Cells["A3"].Value = 11;
		sheet1.Calculate();
		Assert.AreEqual(2, sheet1.Cells["A1"].Value);
	}

	[TestMethod]
	public void RowShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
	{
		var result = new Row().Execute([], CreateContextInScope("A2"));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void RowShouldReturnRowSuppliedAddress()
	{
		var result = new Row().Execute(FunctionsHelper.CreateArgs("A3"), CreateContextWithFakeProvider());
		Assert.AreEqual(3, result.Result);
	}

	[TestMethod]
	public void ColumnShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
	{
		var result = new Column().Execute([], CreateContextInScope("B2"));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void ColumnShouldReturnRowSuppliedAddress()
	{
		var result = new Column().Execute(FunctionsHelper.CreateArgs("E3"), CreateContextWithFakeProvider());
		Assert.AreEqual(5, result.Result);
	}

	[TestMethod]
	public void RowsShouldReturnNbrOfRowsSuppliedRange()
	{
		var result = new Rows().Execute(FunctionsHelper.CreateArgs("A1:B3"), CreateContextWithFakeProvider());
		Assert.AreEqual(3, result.Result);
	}

	[TestMethod]
	public void RowsShouldReturnNbrOfRowsForEntireColumn()
	{
		var result = new Rows().Execute(FunctionsHelper.CreateArgs("A:B"), CreateContextWithFakeProvider());
		Assert.AreEqual(1048576, result.Result);
	}

	[TestMethod]
	public void ColumnssShouldReturnNbrOfRowsSuppliedRange()
	{
		var result = new Columns().Execute(FunctionsHelper.CreateArgs("A1:E3"), CreateContextWithFakeProvider());
		Assert.AreEqual(5, result.Result);
	}

	[TestMethod]
	public void ChooseShouldReturnItemByIndex()
	{
		var func = new Choose();
		var parsingContext = ParsingContext.Create();
		var result = func.Execute(FunctionsHelper.CreateArgs(1, "A", "B"), parsingContext);
		Assert.AreEqual("A", result.Result);
	}

	[TestMethod]
	public void AddressShouldReturnAddressByIndexWithDefaultRefType()
	{
		var result = new AddressFunction().Execute(FunctionsHelper.CreateArgs(1, 2), CreateAddressContext());
		Assert.AreEqual("$B$1", result.Result);
	}

	[TestMethod]
	public void AddressShouldReturnAddressByIndexWithRelativeType()
	{
		var result = new AddressFunction().Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn), CreateAddressContext());
		Assert.AreEqual("B1", result.Result);
	}

	[TestMethod]
	public void AddressShouldReturnAddressByWithSpecifiedWorksheet()
	{
		var result = new AddressFunction().Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, true, "Worksheet1"), CreateAddressContext());
		Assert.AreEqual("Worksheet1!B1", result.Result);
	}

	[TestMethod]
	public void AddressShouldThrowIfR1C1FormatIsSpecified()
	{
		var func = new AddressFunction();
		var parsingContext = CreateAddressContext();
		Assert.ThrowsExactly<InvalidOperationException>(() => func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, false), parsingContext));
	}
}
