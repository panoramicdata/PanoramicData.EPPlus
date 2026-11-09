using OfficeOpenXml;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing;

[TestClass]
public partial class FormulaR1C1Tests
{
	private ExcelPackage _package;
	private ExcelWorksheet _sheet;
	private ExcelWorksheet _sheet2;
	[TestInitialize]
	public void Initialize()
	{
		_package = new ExcelPackage();
		var s1 = ThreeColumnSheet.Add(_package);
		_sheet = s1;
		_sheet2 = _package.Workbook.Worksheets.Add("test2", s1);
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();

	[TestMethod]
	public void RC2()
	{
		var fR1C1 = "RC2";
		_sheet.Cells[5, 1].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 1].Formula;
		_sheet.Cells[5, 1].Formula = f;
		Assert.AreEqual(fR1C1, _sheet.Cells[5, 1].FormulaR1C1);
	}
	[TestMethod]
	public void C()
	{
		var fR1C1 = "SUMIFS(C,C2,RC1)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		_sheet.Cells[5, 3].Formula = f;
		Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
	}
	[TestMethod]
	public void C2Abs()
	{
		var fR1C1 = "SUM(C2)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		Assert.AreEqual("SUM($B:$B)", f);
	}
	[TestMethod]
	public void C2AbsWithSheet()
	{
		var fR1C1 = "SUM(A!C2)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		Assert.AreEqual("SUM(A!$B:$B)", f);
	}
	[TestMethod]
	public void C2()
	{
		var fR1C1 = "SUM(C2)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		_sheet.Cells[5, 3].Formula = f;
		Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
	}
	[TestMethod]
	public void R2Abs()
	{
		var fR1C1 = "SUM(R2)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		Assert.AreEqual("SUM($2:$2)", f);

		fR1C1 = "SUM(TEST2!R2)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		f = _sheet.Cells[5, 3].Formula;
		Assert.AreEqual("SUM(TEST2!$2:$2)", f);

	}
	[TestMethod]
	public void R2()
	{
		var fR1C1 = "SUM(R2)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		_sheet.Cells[5, 3].Formula = f;
		Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
	}
}
