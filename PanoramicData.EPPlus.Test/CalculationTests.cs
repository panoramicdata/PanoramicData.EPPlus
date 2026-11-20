using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test;

[DeploymentItem("Workbooks", "targetFolder")]
[TestClass]
public partial class CalculationTests
{
	[TestMethod]
	public void CalulationTestDatatypes()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("Calc1");
		ws.SetValue("A1", (short)1);
		ws.SetValue("A2", (long)2);
		ws.SetValue("A3", (float)3);
		ws.SetValue("A4", (double)4);
		ws.SetValue("A5", (decimal)5);
		ws.SetValue("A6", (byte)6);
		ws.SetValue("A7", null);
		ws.Cells["A10"].Formula = "Sum(A1:A8)";
		ws.Cells["A11"].Formula = "SubTotal(9,A1:A8)";
		ws.Cells["A12"].Formula = "Average(A1:A8)";

		ws.Calculate();
		Assert.AreEqual(21D, ws.Cells["a10"].Value);
		Assert.AreEqual(21D, ws.Cells["a11"].Value);
		Assert.AreEqual(21D / 6, ws.Cells["a12"].Value);
	}
	[TestMethod]
	public void CalculateTest()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("Calc1");

		ws.SetValue("A1", (short)1);
		var v = ws.Calculate("2.5-A1+ABS(-3.0)-SIN(3)");
		Assert.AreEqual(4.3589, Math.Round((double)v, 4));

		ws.Row(1).Hidden = true;
		v = ws.Calculate("subtotal(109,a1:a10)");
		Assert.AreEqual(0D, v);

		v = ws.Calculate("-subtotal(9,a1:a3)");
		Assert.AreEqual(-1D, v);
	}
	[TestMethod]
	public void CalculateTestIsFunctions()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("Calc1");

		ws.SetValue(1, 1, 1.0D);
		ws.SetFormula(1, 2, "isblank(A1:A5)");
		ws.SetFormula(1, 3, "concatenate(a1,a2,a3)");
		ws.SetFormula(1, 4, "Row()");
		ws.SetFormula(1, 5, "Row(a3)");
		ws.Calculate();
	}
	[TestMethod, Ignore]
	public void Calulation4()
	{
		var pck = new ExcelPackage(new FileInfo(GetFormulaTestPath()));
		pck.Workbook.Calculate();
		Assert.AreEqual(490D, pck.Workbook.Worksheets[1].Cells["D5"].Value);
	}

	private static string GetFormulaTestPath()
	{
#if Core
		var dir = AppContext.BaseDirectory;
		dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
		var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
		return Path.Combine(dir, "Workbooks", "FormulaTest.xlsx");
	}
}
