using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test;

public partial class CalculationTests
{
	[TestMethod]
	public void CalcTwiceError()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("CalcTest");
		ws.Names.AddValue("PRICE", 10);
		ws.Names.AddValue("QUANTITY", 11);
		ws.Cells["A1"].Formula = "PRICE*QUANTITY";
		ws.Names.AddFormula("AMOUNT", "PRICE*QUANTITY");

		ws.Names["PRICE"].Value = 30;
		ws.Names["QUANTITY"].Value = 10;

		ws.Calculate();
		Assert.AreEqual(300D, ws.Cells["A1"].Value);
		Assert.AreEqual(300D, ws.Names["AMOUNT"].Value);
		ws.Names["PRICE"].Value = 40;
		ws.Names["QUANTITY"].Value = 20;

		ws.Calculate();
		Assert.AreEqual(800D, ws.Cells["A1"].Value);
		Assert.AreEqual(800D, ws.Names["AMOUNT"].Value);
	}
	[TestMethod]
	public void IfError()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("CalcTest");
		ws.Cells["A1"].Value = "test1";
		ws.Cells["A5"].Value = "test2";
		ws.Cells["A2"].Value = "Sant";
		ws.Cells["A3"].Value = "Falskt";
		ws.Cells["A4"].Formula = "if(A1>=A5,true,A3)";
		ws.Cells["B1"].Formula = "isText(a1)";
		ws.Cells["B2"].Formula = "isText(\"Test\")";
		ws.Cells["B3"].Formula = "isText(1)";
		ws.Cells["B4"].Formula = "isText(true)";
		ws.Cells["c1"].Formula = "mid(a1,4,15)";

		ws.Calculate();
	}
	[TestMethod]
	public void LeftRightFunctionTest()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("CalcTest");
		ws.SetValue("A1", "asdf");
		ws.Cells["A2"].Formula = "Left(A1, 3)";
		ws.Cells["A3"].Formula = "Left(A1, 10)";
		ws.Cells["A4"].Formula = "Right(A1, 3)";
		ws.Cells["A5"].Formula = "Right(A1, 10)";

		ws.Calculate();
		Assert.AreEqual("asd", ws.Cells["A2"].Value);
		Assert.AreEqual("asdf", ws.Cells["A3"].Value);
		Assert.AreEqual("sdf", ws.Cells["A4"].Value);
		Assert.AreEqual("asdf", ws.Cells["A5"].Value);
	}
	[TestMethod]
	public void IfFunctionTest()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("CalcTest");
		ws.SetValue("A1", 123);
		ws.Cells["A2"].Formula = "IF(A1 = 123, 1, -1)";
		ws.Cells["A3"].Formula = "IF(A1 = 1, 1)";
		ws.Cells["A4"].Formula = "IF(A1 = 1, 1, -1)";
		ws.Cells["A5"].Formula = "IF(A1 = 123, 5)";

		ws.Calculate();
		Assert.AreEqual(1d, ws.Cells["A2"].Value);
		Assert.AreEqual(false, ws.Cells["A3"].Value);
		Assert.AreEqual(-1d, ws.Cells["A4"].Value);
		Assert.AreEqual(5d, ws.Cells["A5"].Value);
	}
	[TestMethod]
	public void INTFunctionTest()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("CalcTest");
		var currentDate = DateTime.UtcNow.Date;
		ws.SetValue("A1", currentDate.ToString("MM/dd/yyyy"));
		ws.SetValue("A2", currentDate.Date);
		ws.SetValue("A3", "31.1");
		ws.SetValue("A4", 31.1);
		ws.Cells["A5"].Formula = "INT(A1)";
		ws.Cells["A6"].Formula = "INT(A2)";
		ws.Cells["A7"].Formula = "INT(A3)";
		ws.Cells["A8"].Formula = "INT(A4)";

		ws.Calculate();
		Assert.AreEqual((int)currentDate.ToOADate(), ws.Cells["A5"].Value);
		Assert.AreEqual((int)currentDate.ToOADate(), ws.Cells["A6"].Value);
		Assert.AreEqual(31, ws.Cells["A7"].Value);
		Assert.AreEqual(31, ws.Cells["A8"].Value);
	}
}
