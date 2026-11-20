using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public class DatabaseTests
{
	private static object CalculateOnThreeColumnDatabase(object a3, object b3, string formula)
	{
		using var package = new ExcelPackage();
		var sheet = package.Workbook.Worksheets.Add("test");
		// database
		sheet.Cells["A1"].Value = "crit1";
		sheet.Cells["B1"].Value = "crit2";
		sheet.Cells["C1"].Value = "crit3";
		sheet.Cells["A2"].Value = "test";
		sheet.Cells["B2"].Value = 2;
		sheet.Cells["C2"].Value = "output";
		sheet.Cells["A3"].Value = a3;
		sheet.Cells["B3"].Value = b3;
		sheet.Cells["C3"].Value = "aaa";
		// criteria
		sheet.Cells["D1"].Value = "crit1";
		sheet.Cells["D2"].Value = "t*t";
		sheet.Cells["E1"].Value = "crit2";
		sheet.Cells["E2"].Value = 2;
		sheet.Cells["F1"].Formula = formula;

		sheet.Workbook.Calculate();

		return sheet.Cells["F1"].Value;
	}

	private static object CalculateOnTwoColumnDatabase(string formula)
	{
		using var package = new ExcelPackage();
		var sheet = package.Workbook.Worksheets.Add("test");
		// database
		sheet.Cells["A1"].Value = "crit1";
		sheet.Cells["B1"].Value = "crit2";

		sheet.Cells["A2"].Value = "test";
		sheet.Cells["B2"].Value = 2;

		sheet.Cells["A3"].Value = "tesst";
		sheet.Cells["B3"].Value = 1;
		// criteria
		sheet.Cells["D1"].Value = "crit1";
		sheet.Cells["D2"].Value = "t*t";
		sheet.Cells["F1"].Formula = formula;

		sheet.Workbook.Calculate();

		return sheet.Cells["F1"].Value;
	}

	[TestMethod]
	public void DgetShouldReturnCorrectResult()
		=> Assert.AreEqual("output", CalculateOnThreeColumnDatabase("test", 3, "DGET(A1:C3,\"Crit3\",D1:E2)"));

	[TestMethod]
	public void DcountShouldReturnCorrectResult()
		=> Assert.AreEqual(1, CalculateOnThreeColumnDatabase("tesst", "2", "DCOUNT(A1:C3,\"Crit2\",D1:E2)"));

	[TestMethod]
	public void DcountaShouldReturnCorrectResult()
		=> Assert.AreEqual(2, CalculateOnThreeColumnDatabase("tesst", "2", "DCOUNTA(A1:C3,\"Crit2\",D1:E2)"));

	[TestMethod]
	public void DMaxShouldReturnCorrectResult()
		=> Assert.AreEqual(2d, CalculateOnTwoColumnDatabase("DMAX(A1:B3,\"Crit2\",D1:E2)"));

	[TestMethod]
	public void DMinShouldReturnCorrectResult()
		=> Assert.AreEqual(1d, CalculateOnTwoColumnDatabase("DMIN(A1:B3,\"Crit2\",D1:E2)"));

	[TestMethod]
	public void DSumShouldReturnCorrectResult()
		=> Assert.AreEqual(3d, CalculateOnTwoColumnDatabase("DSUM(A1:B3,\"Crit2\",D1:E2)"));

	[TestMethod]
	public void DAverageShouldReturnCorrectResult()
		=> Assert.AreEqual(1.5d, CalculateOnTwoColumnDatabase("DAVERAGE(A1:B3,\"Crit2\",D1:E2)"));

	[TestMethod]
	public void DVarShouldReturnCorrectResult()
		=> Assert.AreEqual(0.5d, CalculateOnTwoColumnDatabase("DVAR(A1:B3,\"Crit2\",D1:E2)"));

	[TestMethod]
	public void DVarpShouldReturnCorrectResult()
		=> Assert.AreEqual(0.25d, CalculateOnTwoColumnDatabase("DVARP(A1:B3,\"Crit2\",D1:E2)"));

	[TestMethod]
	public void DVarpShouldReturnByFieldIndex()
		=> Assert.AreEqual(0.25d, CalculateOnTwoColumnDatabase("DVARP(A1:B3,2,D1:E2)"));
}