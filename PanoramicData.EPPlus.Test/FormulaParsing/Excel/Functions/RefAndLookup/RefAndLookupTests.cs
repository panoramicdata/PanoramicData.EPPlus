using OfficeOpenXml;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

[TestClass]
public partial class RefAndLookupTests
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
		=> SheetFormulaHelper.CalculateOnNewSheet("test", "D1", formula, cells);
}
