using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using AddressFunction = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Address;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

public partial class RefAndLookupTests
{
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
