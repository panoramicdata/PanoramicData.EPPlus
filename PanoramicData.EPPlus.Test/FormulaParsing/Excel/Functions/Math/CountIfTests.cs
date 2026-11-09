using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class CountIfTests
{
	private ExcelPackage _package;
	private EpplusExcelDataProvider _provider;
	private ParsingContext _parsingContext;
	private ExcelWorksheet _worksheet;
	private CriteriaRangeHelper _criteria;

	[TestInitialize]
	public void Initialize()
	{
		_package = new ExcelPackage();
		_provider = new EpplusExcelDataProvider(_package);
		_parsingContext = ParsingContext.Create();
		_parsingContext.Scopes.NewScope(RangeAddress.Empty);
		_worksheet = _package.Workbook.Worksheets.Add("testsheet");
		_criteria = new CriteriaRangeHelper(_provider, _worksheet, _parsingContext);
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();

	[TestMethod]
	public void CountIfNumeric() => Assert.AreEqual(2d, _criteria.Execute(new CountIf(), 1d, 2d, 3d, ">1"));

	[TestMethod]
	public void CountIfNonNumeric() => Assert.AreEqual(2d, _criteria.Execute(new CountIf(), "Monday", "Tuesday", "Thursday", "T*day"));

	public void CountIfNullExpression()
	{
		_worksheet.Cells["A1"].Value = null;
		_worksheet.Cells["A2"].Value = 1d;
		_worksheet.Cells["A3"].Value = null;
		_worksheet.Cells["B2"].Value = null;
		var func = new CountIf();
		var range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
		var range2 = _provider.GetRange(_worksheet.Name, 2, 2, 2, 2);
		var args = FunctionsHelper.CreateArgs(range1, range2);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(0d, result.Result);
	}

	[TestMethod]
	public void CountIfNumericExpression() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, 1d, "Not Empty", 1d));

	[TestMethod]
	public void CountIfEqualToEmptyString() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", ""));

	[TestMethod]
	public void CountIfNotEqualToNull() => Assert.AreEqual(2d, _criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", "<>"));

	[TestMethod]
	public void CountIfEqualToZero() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, 0d, "0"));

	[TestMethod]
	public void CountIfNotEqualToZero() => Assert.AreEqual(2d, _criteria.Execute(new CountIf(), null, string.Empty, 0d, "<>0"));

	[TestMethod]
	public void CountIfGreaterThanZero() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, 1d, ">0"));

	[TestMethod]
	public void CountIfGreaterThanOrEqualToZero() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, 1d, ">=0"));

	[TestMethod]
	public void CountIfLessThanZero() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, -1d, "<0"));

	[TestMethod]
	public void CountIfLessThanOrEqualToZero() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, -1d, "<=0"));

	[TestMethod]
	public void CountIfLessThanCharacter() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", "<a"));

	[TestMethod]
	public void CountIfLessThanOrEqualToCharacter() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", "<=a"));

	[TestMethod]
	public void CountIfGreaterThanCharacter() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", ">a"));

	[TestMethod]
	public void CountIfGreaterThanOrEqualToCharacter() => Assert.AreEqual(1d, _criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", ">=a"));
}
