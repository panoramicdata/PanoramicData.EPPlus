using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

/// <summary>Per-test package, worksheet and parsing context for the SUMIF, AVERAGEIF and COUNTIF tests.</summary>
public abstract class CriteriaRangeTestBase
{
	private ExcelPackage _package;

	private protected EpplusExcelDataProvider Provider { get; private set; }

	private protected ParsingContext Context { get; private set; }

	private protected ExcelWorksheet Worksheet { get; private set; }

	private protected CriteriaRangeHelper Criteria { get; private set; }

	[TestInitialize]
	public void Initialize()
	{
		_package = new ExcelPackage();
		Provider = new EpplusExcelDataProvider(_package);
		Context = ParsingContext.Create();
		Context.Scopes.NewScope(RangeAddress.Empty);
		Worksheet = _package.Workbook.Worksheets.Add("testsheet");
		Criteria = new CriteriaRangeHelper(Provider, Worksheet, Context);
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();
}
