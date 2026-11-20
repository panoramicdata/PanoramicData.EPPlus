using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	[Ignore]
	[TestMethod]
	public void ReadBug()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\error.xlsx"));
		var fulla = package.Workbook.Worksheets.FirstOrDefault();
		_ = fulla?.Cells["a:a"]
		.Where(t => !string.IsNullOrWhiteSpace(t.Text)).Select(cell => cell.Value.ToString())
		.ToList();
	}
	[Ignore]
	[TestMethod]
	public void CloseProblem()
	{
		ExcelPackage pck = new();
		var ws = pck.Workbook.Worksheets.Add("Manual Receipts");

		ws.Cells["A1"].Value = " SpaceNeedle Manual Receipt Form";

		using (var r = ws.Cells["A1:F1"])
		{
			r.Merge = true;
			r.Style.Font.Name = "Arial";
			r.Style.Font.Size = 18;
			r.Style.Font.Italic = true;
			r.Style.Font.Color.SetColor(Color.DarkRed);
			r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
		}
		ws.Column(1).Width = 17;
		ws.Column(5).Width = 20;


		ws.Cells["A2"].Value = "Date Produced";

		ws.Cells["A2"].Style.Font.Bold = true;
		ws.Cells["B2"].Value = DateTime.Now.ToString("d");
		ws.Cells["D2"].Value = "Quantity";
		ws.Cells["D2"].Style.Font.Bold = true;
		ws.Cells["E2"].Value = "txt";

		ws.Cells["C4"].Value = "Receipt Number";
		ws.Cells["C4"].Style.WrapText = true;
		ws.Cells["C4"].Style.Font.Bold = true;

		var rowNbr = 5;
		for (var entryNbr = 1; entryNbr <= 1; entryNbr += 1)
		{
			ws.Cells["B" + rowNbr].Value = entryNbr;
			ws.Cells["C" + rowNbr].Value = 1 + entryNbr - 1;
			rowNbr += 1;
		}

		pck.SaveAs(new FileInfo(".\\test.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void OpenXlsm()
	{
		ExcelPackage p = new(new FileInfo("c:\\temp\\cs1.xlsx"));
		_ = p.Workbook.Worksheets.Count;
		p.Save();
	}
	[Ignore]
	[TestMethod]
	public void OpenProblem()
	{
		var xlPackage = new ExcelPackage();
		var ws = xlPackage.Workbook.Worksheets.Add("W1");
		xlPackage.Workbook.Worksheets.Add("W2");

		ws.Cells["A1:A10"].Formula = "W2!A1+C1";
		ws.Cells["B1:B10"].FormulaR1C1 = "W2!R1C1+C1";
		xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void FileLockedProblem()
	{
		using ExcelPackage pck = new(new FileInfo(@"c:\temp\url.xlsx"));
		pck.Workbook.Worksheets[1].DeleteRow(1, 1);
		pck.Save();
		pck.Dispose();
	}
	[Ignore]
	[TestMethod]
	public void RunSample0()
	{
		FileInfo newFile = new(@"c:\temp\bug\sample0.xlsx");
		using ExcelPackage package = new(newFile);
		var worksheet = package.Workbook.Worksheets[1];
		worksheet.InsertColumn(1, 1);

		worksheet.Column(1);

		var last = worksheet.Column(6);
		last.Style.Fill.PatternType = ExcelFillStyle.Solid;
		last.Style.Fill.BackgroundColor.SetColor(Color.Blue);
		last.ColumnMax = 7;
		worksheet.InsertColumn(7, 1);

		//save our new workbook and we are done!
		package.Save();
	}
}
