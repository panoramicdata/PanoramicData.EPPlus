using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	[TestMethod]
	public void Text()
	{
		using ExcelPackage p = new();
		var ws = p.Workbook.Worksheets.Add("Sheet1");
		ws.Cells["A1"].Value = new DateTime(2018, 2, 3);
		ws.Cells["A1"].Style.Numberformat.Format = "d";
		Assert.AreEqual("3", ws.Cells["A1"].Text);
		ws.Cells["A1"].Style.Numberformat.Format = "D";
		Assert.AreEqual("3", ws.Cells["A1"].Text);
		ws.Cells["A1"].Style.Numberformat.Format = "M";
		Assert.AreEqual("2", ws.Cells["A1"].Text);
		ws.Cells["A1"].Style.Numberformat.Format = "Y";
		Assert.AreEqual("18", ws.Cells["A1"].Text);
		ws.Cells["A1"].Style.Numberformat.Format = "YY";
		Assert.AreEqual("18", ws.Cells["A1"].Text);
		ws.Cells["A1"].Style.Numberformat.Format = "YYY";
		Assert.AreEqual("2018", ws.Cells["A1"].Text);
	}
	[TestMethod]
	public void Sort()
	{
		using var package = OpenPackage("Sort.xlsx", true);
		var ws = package.Workbook.Worksheets.Add("Sorting");
		AddSortingData(ws, 1, 1);
		ws.Cells["A:C"].Sort(2, true);
		ws.Cells["A:C"].Sort([2, 1], [true, false]);

		AddSortingData(ws, 1, 5);
		ws.Cells["H1:H1000"].Formula = "E1*2";
		ws.Calculate();

		ws.Cells["E1:H1000"].Sort(3);

		ws.Cells["I5"].Value = 8;
		ws.Cells["I6"].Value = 1;
		ws.Cells["I7"].Value = "12";
		ws.Cells["I8"].Value = "Test";
		ws.Cells["I9"].Value = "aa";
		ws.Cells["I7"].AddComment("Comment", "Jan");


		ws.Cells["I:I"].Sort(0, true);
		package.Save();
	}
	[TestMethod]
	public void ShowGridlines()
	{
		using var p = OpenPackage("ShowGridlines.xlsx");
		var ws = p.Workbook.Worksheets.Add("sort");
		Assert.IsTrue(ws.View.ShowGridLines);   //Default true
		ws.View.ShowGridLines = false;
		Assert.IsFalse(ws.View.ShowGridLines);
		ws.View.ShowGridLines = true;
		Assert.IsTrue(ws.View.ShowGridLines);
	}
	private static void AddSortingData(ExcelWorksheet ws, int row, int col)
	{
		var rand = new Random();
		var addr = ExcelCellBase.GetAddress(row, col, row + 999, col);
		ws.Cells[addr].Style.Fill.PatternType = ExcelFillStyle.Solid;
		ws.Cells[addr].Style.Font.Color.SetColor(Color.White);
		for (var r = row; r <= row + 999; r++)
		{
			var v = rand.NextDouble() * row;
			ws.SetValue(r, col, v);
			ws.SetValue(r, col + 1, $"Rad {r} v={v}");
			ws.SetValue(r, col + 2, $"Rad {r} v={v}");
			ws.Cells[r, col].Style.Fill.BackgroundColor.SetColor(r % 3 == 0 ? Color.Red : r % 3 == 1 ? Color.Green : Color.Blue);
			if (r % 10 == 0)
			{
				ws.Cells[r, col + 1, r, col + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
				ws.Cells[r, col + 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
				ws.Cells[r, col + 2].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
			}
		}
	}
}
