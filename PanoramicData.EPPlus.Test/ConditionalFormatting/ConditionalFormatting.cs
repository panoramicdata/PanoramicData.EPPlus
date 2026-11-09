using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System.Drawing;

namespace PanoramicData.EPPlus.Test.ConditionalFormatting;

/// <summary>
/// Test the Conditional Formatting feature
/// </summary>
[TestClass]
public class ConditionalFormatting
{
	public TestContext TestContext { get; set; }
	private static ExcelPackage? _excelPackage;
	#region Additional test attributes
	// You can use the following additional attributes as you write your tests:
	// Use ClassInitialize to run code before running the first test in the class
	[ClassInitialize]
	public static void MyClassInitialize(TestContext testContext)
	{
		if (!Directory.Exists("Test"))
		{
			Directory.CreateDirectory("Test");
		}

		_excelPackage = new ExcelPackage(new FileInfo(Path.Combine("Test", "ConditionalFormatting.xlsx")));
	}

	// Use ClassCleanup to run code after all tests in a class have run
	[ClassCleanup]
	public static void MyClassCleanup() => _excelPackage = null;

	#endregion

	/// <summary>
	/// 
	/// </summary>
	[TestMethod]
	public void TwoColorScale()
	{
		var ws = _excelPackage.Workbook.Worksheets.Add("ColorScale");
		ws.ConditionalFormatting.AddTwoColorScale(ws.Cells["A1:A5"]);
		SetRowNumbers(ws, 1, 5);
	}

	/// <summary>
	/// 
	/// </summary>
	[TestMethod]
	[Ignore]
	public void ReadConditionalFormatting()
	{
		var pck = new ExcelPackage(new FileInfo(@"c:\temp\cf.xlsx"));

		var ws = pck.Workbook.Worksheets[1];
		Assert.IsTrue(ws.ConditionalFormatting.Count == 6);
		Assert.IsTrue(ws.ConditionalFormatting[0].Type == eExcelConditionalFormattingRuleType.DataBar);

		ws.ConditionalFormatting.AddEqual(ws.Cells["C3"]);
		ws.Cells["C8:C12"].ConditionalFormatting.AddExpression();
		ws.Cells["d12:D22,H12:H22"].ConditionalFormatting.AddFourIconSet(eExcelconditionalFormatting4IconsSetType.RedToBlack);
		pck.SaveAs(new FileInfo(@"c:\temp\cf2.xlsx"));
	}
	/// <summary>
	/// 
	/// </summary>
	[TestMethod]
	[Ignore]
	public void ReadConditionalFormattingError()
	{
		var pck = new ExcelPackage(new FileInfo(@"c:\temp\CofCTemplate.xlsx"));

		_ = pck.Workbook.Worksheets[1];
		pck.SaveAs(new FileInfo(@"c:\temp\cf2.xlsx"));
	}
	/// <summary>
	/// 
	/// </summary>
	[TestMethod]
	public void TwoBackColor()
	{
		var ws = _excelPackage.Workbook.Worksheets.Add("TwoBackColor");
		var condition1 = ws.ConditionalFormatting.AddEqual(ws.Cells["A1"]);
		condition1.StopIfTrue = true;
		condition1.Priority = 1;
		condition1.Formula = "TRUE";
		condition1.Style.Fill.BackgroundColor.Color = Color.Green;
		var condition2 = ws.ConditionalFormatting.AddEqual(ws.Cells["A2"]);
		condition2.StopIfTrue = true;
		condition2.Priority = 2;
		condition2.Formula = "FALSE";
		condition2.Style.Fill.BackgroundColor.Color = Color.Red;
	}
	[TestMethod]
	public void Databar()
	{
		var ws = _excelPackage.Workbook.Worksheets.Add("Databar");
		ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
		SetRowNumbers(ws, 1, 5);
	}
	[TestMethod]
	public void DatabarChangingAddressAddsConditionalFormatNodeInSchemaOrder()
	{
		var ws = _excelPackage.Workbook.Worksheets.Add("DatabarAddressing");
		// Ensure there is at least one element that always exists below ConditionalFormatting nodes.   
		ws.HeaderFooter.AlignWithMargins = true;
		var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
		Assert.AreEqual("sheetData", cf.Node.ParentNode.PreviousSibling.LocalName);
		Assert.AreEqual("headerFooter", cf.Node.ParentNode.NextSibling.LocalName);
		cf.Address = new ExcelAddress("C3");
		Assert.AreEqual("sheetData", cf.Node.ParentNode.PreviousSibling.LocalName);
		Assert.AreEqual("headerFooter", cf.Node.ParentNode.NextSibling.LocalName);
	}
	[TestMethod]
	public void IconSet()
	{
		var ws = _excelPackage.Workbook.Worksheets.Add("IconSet");
		ws.ConditionalFormatting.AddThreeIconSet(ws.Cells["A1:A3"], eExcelconditionalFormatting3IconsSetType.Symbols);
		SetRowNumbers(ws, 1, 3);

		var cf4 = ws.ConditionalFormatting.AddFourIconSet(ws.Cells["B1:B4"], eExcelconditionalFormatting4IconsSetType.Rating);
		cf4.Icon1.Type = eExcelConditionalFormattingValueObjectType.Formula;
		cf4.Icon1.Formula = "0";
		cf4.Icon2.Type = eExcelConditionalFormattingValueObjectType.Formula;
		cf4.Icon2.Formula = "1/3";
		cf4.Icon3.Type = eExcelConditionalFormattingValueObjectType.Formula;
		cf4.Icon3.Formula = "2/3";
		SetRowNumbers(ws, 2, 4);

		var cf5 = ws.ConditionalFormatting.AddFiveIconSet(ws.Cells["C1:C5"], eExcelconditionalFormatting5IconsSetType.Quarters);
		cf5.Icon1.Type = eExcelConditionalFormattingValueObjectType.Num;
		cf5.Icon1.Value = 1;
		cf5.Icon2.Type = eExcelConditionalFormattingValueObjectType.Num;
		cf5.Icon2.Value = 2;
		cf5.Icon3.Type = eExcelConditionalFormattingValueObjectType.Num;
		cf5.Icon3.Value = 3;
		cf5.Icon4.Type = eExcelConditionalFormattingValueObjectType.Num;
		cf5.Icon4.Value = 4;
		cf5.Icon5.Type = eExcelConditionalFormattingValueObjectType.Num;
		cf5.Icon5.Value = 5;
		cf5.ShowValue = false;
		cf5.Reverse = true;

		SetRowNumbers(ws, 3, 5);
	}

	private static void SetRowNumbers(ExcelWorksheet ws, int column, int rows)
	{
		for (var row = 1; row <= rows; row++)
		{
			ws.SetValue(row, column, row);
		}
	}
}
