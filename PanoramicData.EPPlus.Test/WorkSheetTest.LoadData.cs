using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Data;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	private void AutoFilter()
	{
		var ws = _pck.Workbook.Worksheets.Add("Autofilter");
		ws.Cells["A1"].Value = "A1";
		ws.Cells["B1"].Value = "B1";
		ws.Cells["C1"].Value = "C1";
		ws.Cells["D1"].Value = "D1";

		ws.Cells["A2"].Value = 1;
		ws.Cells["B2"].Value = 2;
		ws.Cells["C2"].Value = 3;
		ws.Cells["D2"].Value = 4;

		ws.Cells["A1:D2"].AutoFilter = true;
		ws.Cells["A1:D2"].AutoFilter = false;
		ws.Cells["A1:D2"].AutoFilter = true;
		ws.Cells["A1:D5"].AutoFilter = false;
		ws.Cells["A1:D2"].AutoFilter = true;
	}

	private void AddChartSheet()
	{
		_pck.Workbook.Worksheets.AddChart("ChartSheet", eChartType.ColumnClustered);
		foreach (var _n in _pck.Workbook.Names)
		{
			// Enumerating must not throw.
		}
		foreach (var worksheet in _pck.Workbook.Worksheets)
		{
			if (worksheet is not ExcelChartsheet)
			{
				foreach (var d in worksheet.Drawings)
				{
					// Enumerating must not throw.
				}

				foreach (var d in worksheet.Tables)
				{
					// Enumerating must not throw.
				}

				foreach (var d in worksheet.PivotTables)
				{
					// Enumerating must not throw.
				}

				foreach (var d in worksheet.Names)
				{
					// Enumerating must not throw.
				}

				foreach (var d in worksheet.Comments)
				{
					// Enumerating must not throw.
				}

				foreach (var d in worksheet.ConditionalFormatting)
				{
					// Enumerating must not throw.
				}
			}
		}
	}
	//
	// You can use the following additional attributes as you write your tests:
	//
	// Use ClassInitialize to run code before running the first test in the class
	// Use ClassCleanup to run code after all tests in a class have run
	//[Ignore]
	//[TestMethod]
	public void LoadData()
	{
		var ws = _pck.Workbook.Worksheets.Add("newsheet");
		ws.Cells["T19"].Value = new TimeSpan(3, 30, 30);
		ws.Cells["T20"].Value = new TimeSpan(9, 30, 30);
		ws.Cells["T21"].Value = new TimeSpan(15, 30, 30);
		ws.Cells["T22"].Value = new TimeSpan(21, 30, 30);
		ws.Cells["T23"].Value = new TimeSpan(27, 30, 30);
		ws.Cells["T24"].Value = new TimeSpan(33, 30, 30);
		ws.Cells["T19:T24"].Style.Numberformat.Format = "hh:mm:ss";

		SharedTestData.AddDatesAndValues(ws);
		ws.Cells["v19:v24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
		ws.Cells["v19:v24"].Style.Numberformat.Format = @"$#,##0.00_);($#,##0.00)";

		// int value is actually exposed as double; feature or bug??
		ws.Cells["X19"].Value = 210;
		ws.Cells["X20"].Value = 212;
		ws.Cells["X21"].Value = 221;
		ws.Cells["X22"].Value = 123;
		ws.Cells["X23"].Value = 135;
		ws.Cells["X24"].Value = 134;

		ws.Cells["AA18"].Value = " ";

		ws.Cells["AA19"].Value = "210";
		ws.Cells["AA20"].Value = "212";
		ws.Cells["AA21"].Value = "221";
		ws.Cells["AA22"].Value = "123";
		ws.Cells["AA23"].Value = "135";
		ws.Cells["AA24"].Value = "134";

		ws.Cells["Y19"].Value = true;
		ws.Cells["Y20"].Value = true;
		ws.Cells["Y21"].Value = true;
		ws.Cells["Y22"].Value = false;
		ws.Cells["Y23"].Value = false;
		ws.Cells["Y24"].Value = false;

		ws.Cells["Z19"].Value = "Text1";
		ws.Cells["Z20"].Value = "Text2";
		ws.Cells["Z21"].Value = "Text3";
		ws.Cells["Z22"].Value = "Text4";
		ws.Cells["Z23"].Value = "Text5";
		ws.Cells["Z24"].Value = "Text6";

		// add autofilter
		ws.Cells["U19:X24"].AutoFilter = true;
		var pic = ws.Drawings.AddPicture("Pic1", EPPlusTest.Properties.Resources.Test1);
		pic.SetPosition(150, 140);

		ws.Cells["A30"].Value = "Text orientation 45";
		ws.Cells["A30"].Style.TextRotation = 45;
		ws.Cells["B30"].Value = "Text orientation 90";
		ws.Cells["B30"].Style.TextRotation = 90;
		ws.Cells["C30"].Value = "Text orientation 180";
		ws.Cells["C30"].Style.TextRotation = 180;
		ws.Cells["D30"].Value = "Text orientation 38";
		ws.Cells["D30"].Style.TextRotation = 38;
		ws.Cells["D30"].Style.Font.Bold = true;
		ws.Cells["D30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

		ws.Workbook.Names.Add("TestName", ws.Cells["B30:E30"]);
		ws.Workbook.Names["TestName"].Style.Font.Color.SetColor(Color.Red);


		ws.Workbook.Names["TestName"].Offset(1, 0).Value = "Offset test 1";
		ws.Workbook.Names["TestName"].Offset(2, -1, 2, 2).Value = "Offset test 2";

		//Test vertical align
		ws.Cells["E19"].Value = "Subscript";
		ws.Cells["E19"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Subscript;
		ws.Cells["E20"].Value = "Subscript";
		ws.Cells["E20"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
		ws.Cells["E21"].Value = "Superscript";
		ws.Cells["E21"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
		ws.Cells["E21"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.None;


		ws.Cells["E22"].Value = "Indent 2";
		ws.Cells["E22"].Style.Indent = 2;
		ws.Cells["E23"].Value = "Shrink to fit";
		ws.Cells["E23"].Style.ShrinkToFit = true;

		ws.Cells["e24"].Value = "ReadingOrder LeftToRight";
		ws.Cells["e24"].Style.ReadingOrder = ExcelReadingOrder.LeftToRight;
		ws.Cells["e25"].Value = "ReadingOrder RightToLeft";
		ws.Cells["e25"].Style.ReadingOrder = ExcelReadingOrder.RightToLeft;
		ws.Cells["e26"].Value = "ReadingOrder Context";
		ws.Cells["e26"].Style.ReadingOrder = ExcelReadingOrder.ContextDependent;
		ws.Cells["e27"].Value = "Default Readingorder";

		//Underline

		ws.Cells["F1:F7"].Value = "Underlined";
		ws.Cells["F1"].Style.Font.UnderLineType = ExcelUnderLineType.Single;
		ws.Cells["F2"].Style.Font.UnderLineType = ExcelUnderLineType.Double;
		ws.Cells["F3"].Style.Font.UnderLineType = ExcelUnderLineType.SingleAccounting;
		ws.Cells["F4"].Style.Font.UnderLineType = ExcelUnderLineType.DoubleAccounting;
		ws.Cells["F5"].Style.Font.UnderLineType = ExcelUnderLineType.None;
		ws.Cells["F6:F7"].Style.Font.UnderLine = true;
		ws.Cells["F7"].Style.Font.UnderLine = false;

		ws.Cells["E24"].Value = 0;
		Assert.AreEqual("0", ws.Cells["E24"].Text);
		ws.Cells["F7"].Style.Font.UnderLine = false;
		ws.Names.Add("SheetName", ws.Cells["A1:A2"]);
		ws.View.FreezePanes(3, 5);

		foreach (var cell in ws.Cells["A1"])
		{
			Assert.Fail("A1 is not set");
		}

		foreach (var cell in ws.Cells[ws.Dimension.Address])
		{
			System.Diagnostics.Debug.WriteLine(cell.Address);
		}

		////Linq test
		var res = from c in ws.Cells[ws.Dimension.Address] where c.Value != null && c.Value.ToString() == "Offset test 1" select c;

		foreach (var cell in res)
		{
			System.Diagnostics.Debug.WriteLine(cell.Address);
		}

		_pck.Workbook.Properties.Author = "Jan Källman";
		_pck.Workbook.Properties.Category = "Category";
		_pck.Workbook.Properties.Comments = "Comments";
		_pck.Workbook.Properties.Company = "Adventure works";
		_pck.Workbook.Properties.Keywords = "Keywords";
		_pck.Workbook.Properties.Title = "Title";
		_pck.Workbook.Properties.Subject = "Subject";
		_pck.Workbook.Properties.Status = "Status";
		_pck.Workbook.Properties.HyperlinkBase = new Uri("http://serversideexcel.com", UriKind.Absolute);
		_pck.Workbook.Properties.Manager = "Manager";
		_pck.Workbook.Properties.AppVersion = "4.5.0.1";
		_pck.Workbook.Properties.LinksUpToDate = false;
		_pck.Workbook.Properties.HyperlinksChanged = false;
		_pck.Workbook.Properties.SharedDoc = false;
		_pck.Workbook.Properties.ScaleCrop = false;


		_pck.Workbook.Properties.SetCustomPropertyValue("DateTest", new DateTime(2008, 12, 31));
		Console.WriteLine(_pck.Workbook.Properties.GetCustomPropertyValue("DateTest").ToString());
		_pck.Workbook.Properties.SetCustomPropertyValue("Author", "Jan Källman");
		_pck.Workbook.Properties.SetCustomPropertyValue("Count", 1);
		_pck.Workbook.Properties.SetCustomPropertyValue("IsTested", false);
		_pck.Workbook.Properties.SetCustomPropertyValue("LargeNo", 123456789123);
		_pck.Workbook.Properties.SetCustomPropertyValue("Author", 3);
	}
	const int PERF_ROWS = 5000;
	//[Ignore]
	//[TestMethod]
	public void Performance()
	{
		var ws = _pck.Workbook.Worksheets.Add("Perf");
		Console.WriteLine("StartTime {0}", DateTime.Now);

		for (var i = 1; i <= PERF_ROWS; i++)
		{
			ws.Cells[i, 1].Value = string.Format("Row {0}\n.Test new row\"' öäåü", i);
			ws.Cells[i, 2].Value = i;
			ws.Cells[i, 2].Style.WrapText = true;
			ws.Cells[i, 3].Value = DateTime.Now;
			ws.Cells[i, 4].Value = Random.Shared.NextDouble() * 100000;
		}

		ws.Cells[1, 2, PERF_ROWS, 2].Style.Numberformat.Format = "#,##0";
		ws.Cells[1, 3, PERF_ROWS, 3].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
		ws.Cells[1, 4, PERF_ROWS, 4].Style.Numberformat.Format = "#,##0.00";
		ws.Cells[PERF_ROWS + 1, 2].Formula = "SUM(B1:B" + PERF_ROWS.ToString() + ")";
		ws.Column(1).Width = 12;
		ws.Column(2).Width = 8;
		ws.Column(3).Width = 20;
		ws.Column(4).Width = 14;

		ws.Cells["A1:C1"].Merge = true;
		ws.Cells["A2:A5"].Merge = true;
		ws.DeleteRow(1, 1);
		ws.InsertRow(1, 1);
		ws.InsertRow(3, 1);

		ws.DeleteRow(1000, 3, true);
		ws.DeleteRow(2000, 1, true);

		ws.InsertRow(2001, 4);

		ws.InsertRow(2010, 1, 2010);

		ws.InsertRow(20000, 2);

		ws.DeleteRow(20005, 4, false);

		//Single formula
		ws.Cells["H3"].Formula = "B2+B3";
		ws.DeleteRow(2, 1, true);

		//Shared formula
		ws.Cells["H5:H30"].Formula = "B4+B5";
		ws.Cells["H5:H30"].Style.Numberformat.Format = "_(\"$\"* # ##0.00_);_(\"$\"* (# ##0.00);_(\"$\"* \"-\"??_);_(@_)";
		ws.InsertRow(7, 3);
		ws.InsertRow(2, 1);
		ws.DeleteRow(30, 3, true);

		ws.DeleteRow(15, 2, true);
		ws.Cells["a1:B100"].Style.Locked = false;
		ws.Cells["a1:B12"].Style.Hidden = true;
		Console.WriteLine("EndTime {0}", DateTime.Now);
	}
	private static void ApplyPrintAndViewSettings(ExcelWorksheet ws, string printArea)
	{
		ws.PrinterSettings.PrintArea = null;
		ws.PrinterSettings.PrintArea = ws.Cells[printArea];
		ws.PrinterSettings.PrintArea = null;
		ws.Row(15).PageBreak = true;
		ws.Column(3).PageBreak = true;
		ws.View.ShowHeaders = false;
		ws.View.PageBreakView = true;

		ws.Row(200).Height = 50;
		ws.Workbook.CalcMode = ExcelCalcMode.Automatic;
	}
}
