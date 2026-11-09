using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	[TestMethod]
	public void TestTableNameCanNotStartsWithNumber()
	{
		var ws = _pck.Workbook.Worksheets.Add("Table");
		Assert.ThrowsExactly<ArgumentException>(() => ws.Tables.Add(ws.Cells["A1"], "5TestTable"));
	}

	[TestMethod]
	public void TestTableNameCanNotContainWhiteSpaces()
	{
		var ws = _pck.Workbook.Worksheets.Add("Table");
		Assert.ThrowsExactly<ArgumentException>(() => ws.Tables.Add(ws.Cells["A1"], "Test Table"));
	}

	[TestMethod]
	public void TestTableNameCanStartsWithBackSlash()
	{
		var ws = _pck.Workbook.Worksheets.Add("Table");
		ws.Tables.Add(ws.Cells["A1"], "\\TestTable");
	}

	[TestMethod]
	public void TestTableNameCanStartsWithUnderscore()
	{
		var ws = _pck.Workbook.Worksheets.Add("Table");
		ws.Tables.Add(ws.Cells["A1"], "_TestTable");
	}

	[TestMethod]
	public void TableTotalsRowFunctionEscapesSpecialCharactersInColumnName()
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("TotalsFormulaTest");
		ws.Cells["B1"].Value = "Column1";
		ws.Cells["C1"].Value = "[#'Column'2]";
		var tbl = ws.Tables.Add(ws.Cells["B1:C2"], "TestTable");
		tbl.ShowTotal = true;
		tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
		Assert.AreEqual("SUBTOTAL(109,TestTable['['#''Column''2']])", ws.Cells["C3"].Formula);
	}

	//[Ignore]
	//[TestMethod]
	public void TableTest()
	{
		var ws = _pck.Workbook.Worksheets.Add("Table");
		ws.Cells["B1"].Value = 123;
		var tbl = ws.Tables.Add(ws.Cells["B1:P12"], "TestTable");
		tbl.TableStyle = TableStyles.Custom;

		tbl.ShowFirstColumn = true;
		tbl.ShowTotal = true;
		tbl.ShowHeader = true;
		tbl.ShowLastColumn = true;
		tbl.ShowFilter = false;
		Assert.AreEqual(false, tbl.ShowFilter);
		ws.Cells["K2"].Value = 5;
		ws.Cells["J3"].Value = 4;

		tbl.Columns[8].TotalsRowFunction = RowFunctions.Sum;
		tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])", tbl.Columns[9].Name);
		tbl.Columns[14].CalculatedColumnFormula = "TestTable[[#This Row],[123]]+TestTable[[#This Row],[Column2]]";
		ws.Cells["B2"].Value = 1;
		ws.Cells["B3"].Value = 2;
		ws.Cells["B4"].Value = 3;
		ws.Cells["B5"].Value = 4;
		ws.Cells["B6"].Value = 5;
		ws.Cells["B7"].Value = 6;
		ws.Cells["B8"].Value = 7;
		ws.Cells["B9"].Value = 8;
		ws.Cells["B10"].Value = 9;
		ws.Cells["B11"].Value = 10;
		ws.Cells["B12"].Value = 11;
		ws.Cells["C7"].Value = "Table test";
		ws.Cells["C8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
		ws.Cells["C8"].Style.Fill.BackgroundColor.SetColor(Color.Red);

		ws.Tables.Add(ws.Cells["a12:a13"], "");

		tbl = ws.Tables.Add(ws.Cells["C16:Y35"], "");
		tbl.TableStyle = TableStyles.Medium14;
		tbl.ShowFirstColumn = true;
		tbl.ShowLastColumn = true;
		tbl.ShowColumnStripes = true;
		Assert.AreEqual(true, tbl.ShowFilter);
		tbl.Columns[2].Name = "Test Column Name";

		ws.Cells["G50"].Value = "Timespan";
		ws.Cells["G51"].Value = new DateTime(new TimeSpan(1, 1, 10).Ticks);
		ws.Cells["G52"].Value = new DateTime(1899, 12, 30, 2, 3, 10);
		ws.Cells["G53"].Value = new DateTime(1899, 12, 30, 3, 4, 10);
		ws.Cells["G54"].Value = new DateTime(1899, 12, 30, 4, 5, 10);

		ws.Cells["G51:G55"].Style.Numberformat.Format = "HH:MM:SS";
		tbl = ws.Tables.Add(ws.Cells["G50:G54"], "");
		tbl.ShowTotal = true;
		tbl.ShowFilter = false;
		tbl.Columns[0].TotalsRowFunction = RowFunctions.Sum;
	}

	[TestMethod]
	public void TableDeleteTest()
	{
		using var pkg = new ExcelPackage();
		var wb = pkg.Workbook;
		var sheets = new[]
		{
				wb.Worksheets.Add("WorkSheet A"),
				wb.Worksheets.Add("WorkSheet B")
			};
		for (var i = 1; i <= 4; i++)
		{
			var cell = sheets[0].Cells[1, i];
			cell.Value = cell.Address + "_";
			cell = sheets[1].Cells[1, i];
			cell.Value = cell.Address + "_";
		}

		for (var i = 6; i <= 11; i++)
		{
			var cell = sheets[0].Cells[3, i];
			cell.Value = cell.Address + "_";
			cell = sheets[1].Cells[3, i];
			cell.Value = cell.Address + "_";
		}

		var tables = new[]
		{
				sheets[1].Tables.Add(sheets[1].Cells["A1:D73"], "Tablea"),
				sheets[0].Tables.Add(sheets[0].Cells["A1:D73"], "Table2"),
				sheets[1].Tables.Add(sheets[1].Cells["F3:K10"], "Tableb"),
				sheets[0].Tables.Add(sheets[0].Cells["F3:K10"], "Table3"),
			};
		Assert.AreEqual(5, wb._nextTableID);
		Assert.AreEqual(1, tables[0].Id);
		Assert.AreEqual(2, tables[1].Id);
		try
		{
			sheets[0].Tables.Delete("Tablea");
			Assert.Fail("ArgumentException should have been thrown.");
		}
		catch (ArgumentOutOfRangeException) { /* expected */ }

		sheets[1].Tables.Delete("Tablea");
		Assert.AreEqual(1, tables[1].Id);
		Assert.AreEqual(2, tables[2].Id);

		try
		{
			sheets[1].Tables.Delete(4);
			Assert.Fail("ArgumentException should have been thrown.");
		}
		catch (ArgumentOutOfRangeException)
		{
			// Expected: there is no fifth table.
		}

		var range = sheets[0].Cells[sheets[0].Tables[1].Address.Address];
		sheets[0].Tables.Delete(1, true);
		foreach (var cell in range)
		{
			Assert.IsNull(cell.Value);
		}
	}

	[TestMethod]
	public void TableWithSubtotalsParensInColumnName()
	{
		var ws = _pck.Workbook.Worksheets.Add("Table");
		ws.Cells["B2"].Value = "Header 1";
		ws.Cells["C2"].Value = "Header (2)";
		ws.Cells["B3"].Value = 1;
		ws.Cells["B4"].Value = 2;
		ws.Cells["C3"].Value = 3;
		ws.Cells["C4"].Value = 4;
		var table = ws.Tables.Add(ws.Cells["B2:C4"], "TestTable");
		table.ShowTotal = true;
		table.ShowHeader = true;
		table.Columns[0].TotalsRowFunction = RowFunctions.Sum;
		table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
		ws.Cells["B5"].Calculate();
		Assert.AreEqual(3.0, ws.Cells["B5"].Value);
		ws.Cells["C5"].Calculate();
		Assert.AreEqual(7.0, ws.Cells["C5"].Value);
	}

	//[Ignore]
	//[TestMethod]
	public void CopyTable() => _pck.Workbook.Worksheets.Copy("File4", "Copied table");
}
