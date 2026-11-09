using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;

namespace PanoramicData.EPPlus.Test;

public partial class CellStoreTest
{
	[TestMethod]
	public void DeleteCells()
	{
		var ws = _pck.Workbook.Worksheets.Add("Delete");
		LoadData(ws, 5000);

		ws.DeleteRow(2, 2);
		Assert.AreEqual("3,0", ws.GetValue(2, 1));
		ws.DeleteRow(10, 10);
		Assert.AreEqual("21,0", ws.GetValue(10, 1));
		ws.DeleteRow(50, 40);
		Assert.AreEqual("101,0", ws.GetValue(50, 1));
		ws.DeleteRow(100, 100);
		Assert.AreEqual("251,0", ws.GetValue(100, 1));
		ws.DeleteRow(1, 31);
		Assert.AreEqual("43,0", ws.GetValue(1, 1));
	}
	[TestMethod]
	public void DeleteCellsFirst()
	{
		var ws = _pck.Workbook.Worksheets.Add("DeleteFirst");
		LoadData(ws, 5000);

		ws.DeleteRow(32, 30);
		for (var i = 1; i < 50; i++)
		{
			ws.DeleteRow(1, 1);
		}
	}
	[TestMethod]
	public void DeleteInsert()
	{
		var ws = _pck.Workbook.Worksheets.Add("DeleteInsert");
		LoadData(ws, 5000);

		ws.DeleteRow(2, 33);
		ws.InsertRow(2, 38);

		for (var i = 0; i < 33; i++)
		{
			ws.SetValue(i + 2, 1, i + 2);
		}
	}
	private static void LoadData(ExcelWorksheet ws) => LoadData(ws, 1000);
	private static void LoadData(ExcelWorksheet ws, int rows, int cols = 1, bool isNumeric = false)
	{
		for (var r = 0; r < rows; r++)
		{
			for (var c = 0; c < cols; c++)
			{
				if (isNumeric)
				{
					ws.SetValue(r + 1, c + 1, r + c);
				}
				else
				{
					ws.SetValue(r + 1, c + 1, r.ToString() + "," + c.ToString());
				}
			}
		}
	}
	[TestMethod]
	public void FillInsertTest()
	{
		var ws = _pck.Workbook.Worksheets.Add("FillInsert");

		LoadData(ws, 500);

		var r = 1;
		for (var i = 1; i <= 500; i++)
		{
			ws.InsertRow(r, i);
			Assert.AreEqual((i - 1).ToString() + ",0", ws.GetValue(r + i, 1).ToString());
			r += i + 1;
		}
	}
	[TestMethod]
	public void CopyCellsTest()
	{
		var ws = _pck.Workbook.Worksheets.Add("CopyCells");

		LoadData(ws, 100, isNumeric: true);
		ws.Cells["B1"].Formula = "SUM(A1:A500)";
		ws.Calculate();
		ws.Cells["B1"].Copy(ws.Cells["C1"]);
		ws.Cells["B1"].Copy(ws.Cells["D1"], ExcelRangeCopyOptionFlags.ExcludeFormulas);

		Assert.AreEqual(ws.Cells["B1"].Value, ws.Cells["C1"].Value);
		Assert.AreEqual("SUM(B1:B500)", ws.Cells["C1"].Formula);

		Assert.AreEqual(ws.Cells["B1"].Value, ws.Cells["D1"].Value);
		Assert.AreNotEqual(ws.Cells["B1"].Formula, ws.Cells["D1"].Formula);
	}
	[TestMethod]
	public void Issues351()
	{
		using var package = new ExcelPackage();
		// Arrange
		var worksheet = package.Workbook.Worksheets.Add("Test");
		worksheet.Cells[1, 1].Value = "A";                      // If you remove this "anchor", the problem doesn't happen.
		worksheet.Cells[1026, 1].Value = "B";
		worksheet.Cells[1026, 2].Value = "B";
		var range = worksheet.Row(1026);
		range.Style.Fill.PatternType = ExcelFillStyle.Solid;
		range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0));

		// Act - This should shift the whole row 1026 down 1
		worksheet.InsertRow(1024, 1);

		// Assert - This value should be null, instead it's "B"
		Assert.IsNull(worksheet.Cells[1025, 1].Value);
	}
}
