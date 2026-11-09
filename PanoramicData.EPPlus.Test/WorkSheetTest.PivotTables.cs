using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	[TestMethod]
	public void PivotTableTest()
	{
		_pck = new ExcelPackage();
		var ws = _pck.Workbook.Worksheets.Add("PivotTable");
		ws.Cells["A1"].LoadFromArrays([["A&B", "B\"", "C'", "<D>"]]);
		ws.Cells["A2"].LoadFromArrays(
		[
			[0, 1, 2, 3, 4, 5, 6, 7, 8, 9],
			[9, 8, 7 ,6, 5, 4, 3, 2, 1, 0],
			[1, 1, 2, 3, 5, 8, 13, 21, 34, 55]
		]);
		ws.Tables.Add(ws.Cells["A1:D4"], "PivotData");
		var pt = ws.PivotTables.Add(ws.Cells["G20"], ws.Cells["A1:D4"], "PivotTable1");
		pt.ColumnFields.Add(pt.Fields[1]);
		pt.DataFields.Add(pt.Fields[3]);
		Assert.AreEqual("PivotStyleMedium9", ws.PivotTables["PivotTable1"].StyleName);

		_pck.Workbook.Worksheets.AddChart("PivotChartWorksheet", eChartType.Line, pt);
		SaveWorksheet("Pivot.xlsx");
	}
	[Ignore]
	[TestMethod]
	public void CopyPivotTable()
	{
		_pck.Workbook.Worksheets.Copy("Pivot-Group Date", "Copied Pivottable 1");
		_pck.Workbook.Worksheets.Copy("Pivot-Group Number", "Copied Pivottable 2");
	}
	//[Ignore]
	//[TestMethod]
	public void CreatePivotTable()
	{
		var wsPivot1 = _pck.Workbook.Worksheets.Add("Rows-Data on columns");
		var wsPivot2 = _pck.Workbook.Worksheets.Add("Rows-Data on rows");
		var wsPivot3 = _pck.Workbook.Worksheets.Add("Columns-Data on columns");
		var wsPivot4 = _pck.Workbook.Worksheets.Add("Columns-Data on rows");
		var wsPivot5 = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on columns");
		var wsPivot6 = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on rows");
		var wsPivot7 = _pck.Workbook.Worksheets.Add("Rows/Page-Data on Columns");
		var wsPivot8 = _pck.Workbook.Worksheets.Add("Pivot-Group Date");
		var wsPivot9 = _pck.Workbook.Worksheets.Add("Pivot-Group Number");
		var wsPivot10 = _pck.Workbook.Worksheets.Add("Pivot-Many RowFields");

		var ws = _pck.Workbook.Worksheets.Add("Data");
		ws.Cells["K1"].Value = "Item";
		ws.Cells["L1"].Value = "Category";
		ws.Cells["M1"].Value = "Stock";
		ws.Cells["N1"].Value = "Price";
		ws.Cells["O1"].Value = "Date for grouping";

		ws.Cells["K2"].Value = "Crowbar";
		ws.Cells["L2"].Value = "Hardware";
		ws.Cells["M2"].Value = 12;
		ws.Cells["N2"].Value = 85.2;
		ws.Cells["O2"].Value = new DateTime(2010, 1, 31);

		ws.Cells["K3"].Value = "Crowbar";
		ws.Cells["L3"].Value = "Hardware";
		ws.Cells["M3"].Value = 15;
		ws.Cells["N3"].Value = 12.2;
		ws.Cells["O3"].Value = new DateTime(2010, 2, 28);

		ws.Cells["K4"].Value = "Hammer";
		ws.Cells["L4"].Value = "Hardware";
		ws.Cells["M4"].Value = 550;
		ws.Cells["N4"].Value = 72.7;
		ws.Cells["O4"].Value = new DateTime(2010, 3, 31);

		ws.Cells["K5"].Value = "Hammer";
		ws.Cells["L5"].Value = "Hardware";
		ws.Cells["M5"].Value = 120;
		ws.Cells["N5"].Value = 11.3;
		ws.Cells["O5"].Value = new DateTime(2010, 4, 30);

		ws.Cells["K6"].Value = "Crowbar";
		ws.Cells["L6"].Value = "Hardware";
		ws.Cells["M6"].Value = 120;
		ws.Cells["N6"].Value = 173.2;
		ws.Cells["O6"].Value = new DateTime(2010, 5, 31);

		ws.Cells["K7"].Value = "Hammer";
		ws.Cells["L7"].Value = "Hardware";
		ws.Cells["M7"].Value = 1;
		ws.Cells["N7"].Value = 4.2;
		ws.Cells["O7"].Value = new DateTime(2010, 6, 30);

		ws.Cells["K8"].Value = "Saw";
		ws.Cells["L8"].Value = "Hardware";
		ws.Cells["M8"].Value = 4;
		ws.Cells["N8"].Value = 33.12;
		ws.Cells["O8"].Value = new DateTime(2010, 6, 28);

		ws.Cells["K9"].Value = "Screwdriver";
		ws.Cells["L9"].Value = "Hardware";
		ws.Cells["M9"].Value = 1200;
		ws.Cells["N9"].Value = 45.2;
		ws.Cells["O9"].Value = new DateTime(2010, 8, 31);

		ws.Cells["K10"].Value = "Apple";
		ws.Cells["L10"].Value = "Groceries";
		ws.Cells["M10"].Value = 807;
		ws.Cells["N10"].Value = 1.2;
		ws.Cells["O10"].Value = new DateTime(2010, 9, 30);

		ws.Cells["K11"].Value = "Butter";
		ws.Cells["L11"].Value = "Groceries";
		ws.Cells["M11"].Value = 52;
		ws.Cells["N11"].Value = 7.2;
		ws.Cells["O11"].Value = new DateTime(2010, 10, 31);
		ws.Cells["O2:O11"].Style.Numberformat.Format = "yyyy-MM-dd";

		var pt = wsPivot1.PivotTables.Add(wsPivot1.Cells["A1"], ws.Cells["K1:N11"], "Pivottable1");
		pt.GrandTotalCaption = "Total amount";
		pt.RowFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[0]);
		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataFields[0].Function = DataFieldFunctions.Product;
		pt.DataOnRows = false;

		pt = wsPivot2.PivotTables.Add(wsPivot2.Cells["A1"], ws.Cells["K1:N11"], "Pivottable2");
		pt.RowFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[0]);
		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataFields[0].Function = DataFieldFunctions.Average;
		pt.DataOnRows = true;

		pt = wsPivot3.PivotTables.Add(wsPivot3.Cells["A1"], ws.Cells["K1:N11"], "Pivottable3");
		pt.ColumnFields.Add(pt.Fields[1]);
		pt.ColumnFields.Add(pt.Fields[0]);
		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataOnRows = false;

		pt = wsPivot4.PivotTables.Add(wsPivot4.Cells["A1"], ws.Cells["K1:N11"], "Pivottable4");
		pt.ColumnFields.Add(pt.Fields[1]);
		pt.ColumnFields.Add(pt.Fields[0]);
		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataOnRows = true;

		pt = wsPivot5.PivotTables.Add(wsPivot5.Cells["A1"], ws.Cells["K1:N11"], "Pivottable5");
		pt.ColumnFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[0]);
		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataOnRows = false;

		pt = wsPivot6.PivotTables.Add(wsPivot6.Cells["A1"], ws.Cells["K1:N11"], "Pivottable6");
		pt.ColumnFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[0]);
		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataOnRows = true;
		wsPivot6.Drawings.AddChart("Pivotchart6", eChartType.BarStacked3D, pt);

		pt = wsPivot7.PivotTables.Add(wsPivot7.Cells["A3"], ws.Cells["K1:N11"], "Pivottable7");
		pt.PageFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[0]);
		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataOnRows = false;

		pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Max;
		Assert.AreEqual(eSubTotalFunctions.Sum | eSubTotalFunctions.Max, pt.Fields[0].SubTotalFunctions);

		pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP;
		Assert.AreEqual(eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP, pt.Fields[0].SubTotalFunctions);

		pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.None;
		Assert.AreEqual(eSubTotalFunctions.None, pt.Fields[0].SubTotalFunctions);

		pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Default;
		Assert.AreEqual(eSubTotalFunctions.Default, pt.Fields[0].SubTotalFunctions);

		pt.Fields[0].Sort = eSortType.Descending;
		pt.TableStyle = TableStyles.Medium14;

		pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["A3"], ws.Cells["K1:O11"], "Pivottable8");
		pt.RowFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[4]);
		pt.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days | eDateGroupBy.Quarters, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
		pt.RowHeaderCaption = "År";
		pt.Fields[4].Name = "Dag";
		pt.Fields[5].Name = "Månad";
		pt.Fields[6].Name = "Kvartal";
		pt.GrandTotalCaption = "Totalt";

		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataOnRows = true;

		pt = wsPivot9.PivotTables.Add(wsPivot9.Cells["A3"], ws.Cells["K1:N11"], "Pivottable9");
		pt.PageFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[3]);
		pt.RowFields[0].AddNumericGrouping(-3.3, 5.5, 4.0);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataOnRows = false;
		pt.TableStyle = TableStyles.Medium14;

		pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["H3"], ws.Cells["K1:O11"], "Pivottable10");
		pt.RowFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[4]);
		pt.Fields[4].AddDateGrouping(7, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
		pt.RowHeaderCaption = "Veckor";
		pt.GrandTotalCaption = "Totalt";

		pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["A60"], ws.Cells["K1:O11"], "Pivottable11");
		pt.RowFields.Add(pt.Fields["Category"]);
		pt.RowFields.Add(pt.Fields["Item"]);
		pt.RowFields.Add(pt.Fields["Date for grouping"]);

		pt.DataFields.Add(pt.Fields[3]);
		pt.DataFields.Add(pt.Fields[2]);
		pt.DataOnRows = true;

		pt = wsPivot10.PivotTables.Add(wsPivot10.Cells["A1"], ws.Cells["K1:O11"], "Pivottable10");
		pt.ColumnFields.Add(pt.Fields[1]);
		pt.RowFields.Add(pt.Fields[0]);
		pt.RowFields.Add(pt.Fields[3]);
		pt.RowFields.Add(pt.Fields[2]);
		pt.RowFields.Add(pt.Fields[4]);
		pt.DataOnRows = true;
		pt.ColumnHeaderCaption = "Column Caption";
		pt.RowHeaderCaption = "Row Caption";
	}
	[Ignore]
	[TestMethod]
	public void ReadPivotTable()
	{
		ExcelPackage pck = new(new FileInfo(@"c:\temp\pivot\pivotforread.xlsx"));

		var pivot1 = pck.Workbook.Worksheets[2].PivotTables[0];

		Assert.AreEqual(24, pivot1.Fields.Count);
		Assert.AreEqual(3, pivot1.RowFields.Count);
		Assert.AreEqual(7, pivot1.DataFields.Count);
		Assert.AreEqual(0, pivot1.ColumnFields.Count);

		Assert.AreEqual("Sum of n3", pivot1.DataFields[1].Name);
		Assert.AreEqual(eSortType.Ascending, pivot1.Fields[2].Sort);

		Assert.AreEqual(false, pivot1.DataOnRows);

		_ = pck.Workbook.Worksheets[2].PivotTables[0];
		_ = pck.Workbook.Worksheets[3].PivotTables[0];

		_ = pck.Workbook.Worksheets[4].PivotTables[0];
		var pivot5 = pck.Workbook.Worksheets[5].PivotTables[0];
		pivot5.CacheDefinition.SourceRange = pck.Workbook.Worksheets[1].Cells["Q1:X300"];

		var pivot6 = pck.Workbook.Worksheets[6].PivotTables[0];

		pck.Workbook.Worksheets[6].Drawings.AddChart("chart1", eChartType.ColumnStacked3D, pivot6);

		pck.SaveAs(new FileInfo(@"c:\temp\pivot\pivotforread_new.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void CreatePivotMultData()
	{
		FileInfo fi = new(@"c:\temp\test.xlsx");
		ExcelPackage pck = new(fi);

		var ws = pck.Workbook.Worksheets.Add("Data");
		var pv = pck.Workbook.Worksheets.Add("Pivot");

		ws.Cells["A1"].Value = "Data1";
		ws.Cells["B1"].Value = "Data2";

		ws.Cells["A2"].Value = "1";
		ws.Cells["B2"].Value = "2";

		ws.Cells["A3"].Value = "3";
		ws.Cells["B3"].Value = "4";

		ws.Select("A1:B3");

		var pt = pv.PivotTables.Add(pv.SelectedRange, ws.SelectedRange, "Pivot");

		pt.RowFields.Add(pt.Fields["Data2"]);

		var df = pt.DataFields.Add(pt.Fields["Data1"]);
		df.Function = DataFieldFunctions.Count;

		df = pt.DataFields.Add(pt.Fields["Data1"]);
		df.Function = DataFieldFunctions.Sum;

		df = pt.DataFields.Add(pt.Fields["Data1"]);
		df.Function = DataFieldFunctions.StdDev;
		df.Name = "DatA1_2";

		pck.Save();
	}
}
