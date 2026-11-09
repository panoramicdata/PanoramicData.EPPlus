using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[TestMethod, Ignore]
	public void Issue15198()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\Output.xlsx"), new FileInfo(@"c:\temp\bug\demo.xlsx"));
		var workBook = package.Workbook;
		var worksheet = workBook.Worksheets[1];

		worksheet.DeleteRow(12);

		package.Save();
	}
	[TestMethod, Ignore]
	public void Issue13492()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\Bug13492.xlsx"));
		var workBook = package.Workbook;
		var worksheet = workBook.Worksheets[1];

		_ = worksheet.Cells["K31"].RichText.Text;

		package.Save();
	}
	[TestMethod, Ignore]
	public void Issue14966()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\ssis\FileFromReportingServer2012.xlsx"));
		package.SaveAs(new FileInfo(@"c:\temp\bug\ssis\Corrupted.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue15200()
	{
		File.Copy(@"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\input.xlsx", @"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\output.xlsx", true);

		using var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\output.xlsx"));
		var sheet = p.Workbook.Worksheets.First();

		var sourceRange = sheet.Cells[1, 1, 1, 2];
		var resultRange = sheet.Cells[3, 1, 3, 2];
		sourceRange.Copy(resultRange);

		sourceRange = sheet.Cells[1, 1, 1, 7];
		resultRange = sheet.Cells[5, 1, 5, 7];
		sourceRange.Copy(resultRange);  // This throws System.ArgumentException: Can't merge and already merged range

		sourceRange = sheet.Cells[1, 1, 1, 7];
		resultRange = sheet.Cells[7, 3, 7, 7];
		sourceRange.Copy(resultRange);  // This throws System.ArgumentException: Can't merge and already merged range

		p.Save();
	}
	[TestMethod]
	public void Issue15212()
	{
		var s = "_(\"R$ \"* #,##0.00_);_(\"R$ \"* (#,##0.00);_(\"R$ \"* \"-\"??_);_(@_) )";
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("StyleBug");
		ws.Cells["A1"].Value = 5698633.64;
		ws.Cells["A1"].Style.Numberformat.Format = s;
		_ = ws.Cells["A1"].Text;
	}
	[TestMethod, Ignore]
	public void Issue15213()
	{
		using var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\ExcelClearDemo\exceltestfile.xlsx"));
		foreach (var ws in p.Workbook.Worksheets)
		{
			ws.Cells[1023, 1, ws.Dimension.End.Row - 2, ws.Dimension.End.Column].Clear();
			Assert.AreNotEqual(null, ws.Dimension);
		}

		foreach (var cell in p.Workbook.Worksheets[2].Cells)
		{
			Console.WriteLine(cell);
		}

		p.SaveAs(new FileInfo(@"c:\temp\bug\ExcelClearDemo\exceltestfile-save.xlsx"));
	}

	[TestMethod, Ignore]
	public void Issuer15217()
	{

		using var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\FormatRowCol.xlsx"));
		var ws = p.Workbook.Worksheets.Add("fmt");
		ws.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
		ws.Row(1).Style.Fill.BackgroundColor.SetColor(Color.LightGray);
		ws.Cells["A1:B2"].Value = 1;
		ws.Column(1).Style.Numberformat.Format = "yyyy-mm-dd hh:mm";
		ws.Column(2).Style.Numberformat.Format = "#,##0";
		p.Save();
	}
	[TestMethod, Ignore]
	public void Issuer15228()
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("colBug");

		var col = ws.Column(7);
		col.ColumnMax = 8;
		col.Hidden = true;

		var col8 = ws.Column(8);
		Assert.AreEqual(true, col8.Hidden);
	}

	[TestMethod, Ignore]
	public void Issue15234()
	{
		using var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\merge2\input.xlsx"));
		var sheet = p.Workbook.Worksheets.First();

		var sourceRange = sheet.Cells["1:4"];

		sheet.InsertRow(5, 4);

		var resultRange = sheet.Cells["5:8"];
		sourceRange.Copy(resultRange);

		p.Save();
	}
	[TestMethod]
	/**** Pivottable issue ****/
	public void Issue()
	{
		FileInfo MyFile = new(@"c:\temp\bug\pivottable.xlsx");
		LoadData(MyFile);
		BuildPivotTable1(MyFile);
		BuildPivotTable2(MyFile);
	}

	private static void LoadData(FileInfo MyFile)
	{
		if (MyFile.Exists)
		{
			MyFile.Delete();  // ensures we create a new workbook
		}

		using ExcelPackage EP = new(MyFile);
		// add a new worksheet to the empty workbook
		var wsData = EP.Workbook.Worksheets.Add("Data");
		//Add the headers
		wsData.Cells[1, 1].Value = "INVOICE_DATE";
		wsData.Cells[1, 2].Value = "TOTAL_INVOICE_PRICE";
		wsData.Cells[1, 3].Value = "EXTENDED_PRICE_VARIANCE";
		wsData.Cells[1, 4].Value = "AUDIT_LINE_STATUS";
		wsData.Cells[1, 5].Value = "RESOLUTION_STATUS";
		wsData.Cells[1, 6].Value = "COUNT";

		//Add some items...
		wsData.Cells["A2"].Value = Convert.ToDateTime("04/2/2012");
		wsData.Cells["B2"].Value = 33.63;
		wsData.Cells["C2"].Value = -.87;
		wsData.Cells["D2"].Value = "Unfavorable Price Variance";
		wsData.Cells["E2"].Value = "Pending";
		wsData.Cells["F2"].Value = 1;

		wsData.Cells["A3"].Value = Convert.ToDateTime("04/2/2012");
		wsData.Cells["B3"].Value = 43.14;
		wsData.Cells["C3"].Value = -1.29;
		wsData.Cells["D3"].Value = "Unfavorable Price Variance";
		wsData.Cells["E3"].Value = "Pending";
		wsData.Cells["F3"].Value = 1;

		wsData.Cells["A4"].Value = Convert.ToDateTime("11/8/2011");
		wsData.Cells["B4"].Value = 55;
		wsData.Cells["C4"].Value = -2.87;
		wsData.Cells["D4"].Value = "Unfavorable Price Variance";
		wsData.Cells["E4"].Value = "Pending";
		wsData.Cells["F4"].Value = 1;

		wsData.Cells["A5"].Value = Convert.ToDateTime("11/8/2011");
		wsData.Cells["B5"].Value = 38.72;
		wsData.Cells["C5"].Value = -5.00;
		wsData.Cells["D5"].Value = "Unfavorable Price Variance";
		wsData.Cells["E5"].Value = "Pending";
		wsData.Cells["F5"].Value = 1;

		wsData.Cells["A6"].Value = Convert.ToDateTime("3/4/2011");
		wsData.Cells["B6"].Value = 77.44;
		wsData.Cells["C6"].Value = -1.55;
		wsData.Cells["D6"].Value = "Unfavorable Price Variance";
		wsData.Cells["E6"].Value = "Pending";
		wsData.Cells["F6"].Value = 1;

		wsData.Cells["A7"].Value = Convert.ToDateTime("3/4/2011");
		wsData.Cells["B7"].Value = 127.55;
		wsData.Cells["C7"].Value = -10.50;
		wsData.Cells["D7"].Value = "Unfavorable Price Variance";
		wsData.Cells["E7"].Value = "Pending";
		wsData.Cells["F7"].Value = 1;

		using (var range = wsData.Cells[2, 1, 7, 1])
		{
			range.Style.Numberformat.Format = "mm-dd-yy";
		}

		wsData.Cells.AutoFitColumns(0);
		EP.Save();
	}
	private static void BuildPivotTable1(FileInfo MyFile)
	{
		using ExcelPackage ep = new(MyFile);

		var wsData = ep.Workbook.Worksheets["Data"];
		var totalRows = wsData.Dimension.Address;
		var data = wsData.Cells[totalRows];

		var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot1");

		var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit1");
		pivotTable1.ColumnGrandTotals = true;
		var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


		rowField.AddDateGrouping(eDateGroupBy.Years);
		var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
		yearField.Name = "Year";

		pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

		var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
		TotalSpend.Name = "Total Spend";
		TotalSpend.Format = "$##,##0";


		var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
		CountInvoicePrice.Name = "Total Lines";
		CountInvoicePrice.Format = "##,##0";

		pivotTable1.DataOnRows = false;
		ep.Save();
		ep.Dispose();

	}

	private static void BuildPivotTable2(FileInfo MyFile)
	{
		using ExcelPackage ep = new(MyFile);

		var wsData = ep.Workbook.Worksheets["Data"];
		var totalRows = wsData.Dimension.Address;
		var data = wsData.Cells[totalRows];

		var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot2");

		var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit2");
		pivotTable1.ColumnGrandTotals = true;
		var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


		rowField.AddDateGrouping(eDateGroupBy.Years);
		var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
		yearField.Name = "Year";

		pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

		var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
		TotalSpend.Name = "Total Spend";
		TotalSpend.Format = "$##,##0";


		var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
		CountInvoicePrice.Name = "Total Lines";
		CountInvoicePrice.Format = "##,##0";

		pivotTable1.DataOnRows = false;
		ep.Save();
		ep.Dispose();

	}

	[TestMethod, Ignore]
	public void issue15249()
	{
		using var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\Boldtextcopy.xlsx"));
		exfile.Workbook.Worksheets.Copy("sheet1", "copiedSheet");
		exfile.SaveAs(new FileInfo(@"c:\temp\bug\Boldtextcopy2.xlsx"));
	}
	[TestMethod, Ignore]
	public void issue15300()
	{
		using var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\headfootpic.xlsx"));
		exfile.Workbook.Worksheets.Copy("sheet1", "copiedSheet");
		exfile.SaveAs(new FileInfo(@"c:\temp\bug\headfootpic_save.xlsx"));

	}
	[TestMethod, Ignore]
	public void issue15295()
	{
		using var exfile = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivot issue\input.xlsx"));
		exfile.SaveAs(new FileInfo(@"C:\temp\bug\pivot issue\pivotcoldup.xlsx"));

	}
}
