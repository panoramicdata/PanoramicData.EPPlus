using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public partial class MathFunctionsTests : FormulaParserTestBase
{
	private ExcelPackage _package;

	[TestInitialize]
	public void Setup()
	{
		_package = new ExcelPackage();
		var excelDataProvider = new EpplusExcelDataProvider(_package);
		_parser = new FormulaParser(excelDataProvider);
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();
}
