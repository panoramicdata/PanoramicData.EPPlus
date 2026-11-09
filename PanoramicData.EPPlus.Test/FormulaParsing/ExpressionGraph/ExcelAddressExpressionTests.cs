using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using FakeItEasy;

namespace PanoramicData.EPPlus.Test.FormulaParsing.ExpressionGraph;

[TestClass]
public class ExcelAddressExpressionTests
{
	private ParsingContext _parsingContext;
	private ParsingScope _scope;

	[TestInitialize]
	public void Setup()
	{
		_parsingContext = ParsingContext.Create();
		_scope = _parsingContext.Scopes.NewScope(RangeAddress.Empty);
	}

	[TestCleanup]
	public void Cleanup() => _scope.Dispose();

	[TestMethod]
	public void ConstructorShouldThrowIfExcelDataProviderIsNull() => Assert.ThrowsExactly<ArgumentNullException>(() => new ExcelAddressExpression("A1", null, _parsingContext));

	[TestMethod]
	public void ConstructorShouldThrowIfParsingContextIsNull() => Assert.ThrowsExactly<ArgumentNullException>(() => new ExcelAddressExpression("A1", A.Fake<ExcelDataProvider>(), null));

	private static void CompileAndAssert(string address, Action<CompileResult> assert, bool resolveAsRange = false, Action<ExcelWorksheet>? arrange = null)
	{
		var parsingContext = ParsingContext.Create();
		var file = new FileInfo("filename.xlsx");
		using var package = new ExcelPackage(file);
		using var sheet = package.Workbook.Worksheets.Add("NewSheet");
		using var excelDataProvider = new EpplusExcelDataProvider(package);
		arrange?.Invoke(sheet);
		var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
		using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
		{
			var expression = new ExcelAddressExpression(address, excelDataProvider, parsingContext);
			if (resolveAsRange)
			{
				expression.ResolveAsRange = true;
			}

			assert(expression.Compile());
		}
	}

	private static ExcelDataProvider.IRangeInfo AssertRange(CompileResult result, string expectedAddress)
	{
		var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
		Assert.IsNotNull(rangeInfo);
		Assert.AreEqual(expectedAddress, rangeInfo.Address.Address);
		return rangeInfo;
	}

	private static void AssertEmptyRange(CompileResult result, string expectedAddress)
	{
		var rangeInfo = AssertRange(result, expectedAddress);
		// Enumerating the range still yields no results.
		Assert.AreEqual(0, rangeInfo.Count());
	}

	#region Compile Tests
	[TestMethod]
	public void CompileSingleCellReference()
		=> CompileAndAssert("A1", result => Assert.IsNull(result.Result));

	[TestMethod]
	public void CompileSingleCellReferenceWithValue()
		=> CompileAndAssert("A1", result => Assert.AreEqual("Value", result.Result), arrange: sheet => sheet.Cells[1, 1].Value = "Value");

	[TestMethod]
	public void CompileSingleCellReferenceResolveToRange()
		=> CompileAndAssert("A1", result => AssertRange(result, "A1"), resolveAsRange: true);

	[TestMethod]
	public void CompileSingleCellReferenceResolveToRangeColumnAbsolute()
		=> CompileAndAssert("$A1", result => AssertRange(result, "$A1"), resolveAsRange: true);

	[TestMethod]
	public void CompileSingleCellReferenceResolveToRangeRowAbsolute()
		=> CompileAndAssert("$A1", result => AssertRange(result, "$A1"), resolveAsRange: true);

	[TestMethod]
	public void CompileSingleCellReferenceResolveToRangeAbsolute()
		=> CompileAndAssert("$A$1", result => AssertRange(result, "$A$1"), resolveAsRange: true);

	[TestMethod]
	public void CompileMultiCellReference()
		=> CompileAndAssert("A1:A5", result => AssertEmptyRange(result, "A1:A5"));

	[TestMethod]
	public void CompileMultiCellReferenceWithValues()
		=> CompileAndAssert("A1:A5", result =>
		{
			var rangeInfo = AssertRange(result, "A1:A5");
			Assert.AreEqual(5, rangeInfo.Count());
			for (var i = 1; i <= 5; i++)
			{
				var rangeItem = rangeInfo.ElementAt(i - 1);
				Assert.AreEqual("Value" + i, rangeItem.Value);
			}
		}, arrange: sheet =>
		{
			sheet.Cells[1, 1].Value = "Value1";
			sheet.Cells[2, 1].Value = "Value2";
			sheet.Cells[3, 1].Value = "Value3";
			sheet.Cells[4, 1].Value = "Value4";
			sheet.Cells[5, 1].Value = "Value5";
		});

	[TestMethod]
	public void CompileMultiCellReferenceColumnAbsolute()
		=> CompileAndAssert("$A1:$A5", result => AssertEmptyRange(result, "$A1:$A5"));

	[TestMethod]
	public void CompileMultiCellReferenceRowAbsolute()
		=> CompileAndAssert("A$1:A$5", result => AssertEmptyRange(result, "A$1:A$5"));

	[TestMethod]
	public void CompileMultiCellReferenceAbsolute()
		=> CompileAndAssert("$A$1:$A$5", result => AssertEmptyRange(result, "$A$1:$A$5"));
	#endregion
}