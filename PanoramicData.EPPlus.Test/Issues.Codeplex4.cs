using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[TestMethod, Ignore]
	public void issue15282()
	{
		using var exfile = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivottable-table.xlsx"));
		exfile.SaveAs(new FileInfo(@"C:\temp\bug\pivot issue\pivottab-tab-save.xlsx"));

	}

	[TestMethod, Ignore]
	public void Issues14699()
	{
		FileInfo newFile = new("c:\\temp\\bug\\EPPlus_Issue14699.xlsx");
		ExcelPackage pkg = new(newFile);
		var wksheet = pkg.Workbook.Worksheets.Add("Issue14699");
		// Initialize a small range
		for (var row = 1; row < 11; row++)
		{
			for (var col = 1; col < 11; col++)
			{
				wksheet.Cells[row, col].Value = string.Format("{0}{1}", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[col - 1], row);
			}
		}

		wksheet.View.FreezePanes(3, 3);
		pkg.Save();

	}
	[TestMethod, Ignore]
	public void Issue15382()
	{
		using var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\Text Run Issue.xlsx"));
		exfile.SaveAs(new FileInfo(@"C:\temp\bug\inlinText.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue15380()
	{
		using var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\dotinname.xlsx"));
		var v = exfile.Workbook.Worksheets["sheet1.3"].Names["Test.Name"].Value;
		Assert.AreEqual(1, v);
	}
	[TestMethod, Ignore]
	public void Issue15378()
	{
		using var p = new ExcelPackage(new FileInfo(@"c:\temp\bubble.xlsx"));
		var c = p.Workbook.Worksheets[1].Drawings[0] as ExcelBubbleChart;
		_ = c.Series[0] as ExcelBubbleChartSerie;
	}
	[TestMethod]
	public void Issue15377()
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("ws1");
		ws.Cells["A1"].Value = (double?)1;
		_ = ws.GetValue<double?>(1, 1);
	}
	[TestMethod]
	public void Issue15374()
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("RT");
		var r = ws.Cells["A1"];
		r.RichText.Text = "Cell 1";
		r["A2"].RichText.Add("Cell 2");
		p.SaveAs(new FileInfo(@"c:\temp\rt.xlsx"));
	}
	[TestMethod]
	public void IssueTranslate()
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("Trans");
		ws.Cells["A1:A2"].Formula = "IF(1=1, \"A's B C\",\"D\") ";
		var fr = ws.Cells["A1:A2"].FormulaR1C1;
		ws.Cells["A1:A2"].FormulaR1C1 = fr;
		Assert.AreEqual("IF(1=1,\"A's B C\",\"D\")", ws.Cells["A2"].Formula);
	}
	[TestMethod]
	public void Issue15397()
	{
		using var p = new ExcelPackage();
		var workSheet = p.Workbook.Worksheets.Add("styleerror");
		workSheet.Cells["F:G"].Style.Fill.PatternType = ExcelFillStyle.Solid;
		workSheet.Cells["F:G"].Style.Fill.BackgroundColor.SetColor(Color.Red);

		workSheet.Cells["A:A,C:C"].Style.Fill.PatternType = ExcelFillStyle.Solid;
		workSheet.Cells["A:A,C:C"].Style.Fill.BackgroundColor.SetColor(Color.Red);

		//And then: 

		workSheet.Cells["A:H"].Style.Font.Color.SetColor(Color.Blue);

		workSheet.Cells["I:I"].Style.Fill.PatternType = ExcelFillStyle.Solid;
		workSheet.Cells["I:I"].Style.Fill.BackgroundColor.SetColor(Color.Red);
		workSheet.Cells["I2"].Style.Fill.BackgroundColor.SetColor(Color.Green);
		workSheet.Cells["I4"].Style.Fill.BackgroundColor.SetColor(Color.Blue);
		workSheet.Cells["I9"].Style.Fill.BackgroundColor.SetColor(Color.Pink);

		workSheet.InsertColumn(2, 2, 9);
		workSheet.Column(45).Width = 0;

		p.SaveAs(new FileInfo(@"c:\temp\styleerror.xlsx"));
	}
	[TestMethod]
	public void Issuer14801()
	{
		using var p = new ExcelPackage();
		var workSheet = p.Workbook.Worksheets.Add("rterror");
		var cell = workSheet.Cells["A1"];
		cell.RichText.Add("toto: ");
		cell.RichText[0].PreserveSpace = true;
		cell.RichText[0].Bold = true;
		cell.RichText.Add("tata");
		cell.RichText[1].Bold = false;
		cell.RichText[1].Color = Color.Green;
		p.SaveAs(new FileInfo(@"c:\temp\rtpreserve.xlsx"));
	}
	[TestMethod]
	public void Issuer15445()
	{
		using var p = new ExcelPackage();
		p.Workbook.Worksheets.Add("ws1");
		var ws2 = p.Workbook.Worksheets.Add("ws2");
		ws2.View.SelectedRange = "A1:B3 D12:D15";
		ws2.View.ActiveCell = "D15";
		p.SaveAs(new FileInfo(@"c:\temp\activeCell.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue15429()
	{
		FileInfo file = new(@"c:\temp\original.xlsx");
		using (ExcelPackage excelPackage = new(file))
		{
			var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
			var equalsRule = worksheet.ConditionalFormatting.AddEqual(new ExcelAddress(2, 3, 6, 3));
			equalsRule.Formula = "0";
			equalsRule.Style.Fill.BackgroundColor.Color = Color.Blue;
			worksheet.ConditionalFormatting.AddDatabar(new ExcelAddress(4, 4, 4, 4), Color.Red);
			excelPackage.Save();
		}

		using (ExcelPackage excelPackage = new(file))
		{
			var worksheet = excelPackage.Workbook.Worksheets["Sheet 1"];
			var i = 0;
			foreach (var conditionalFormat in worksheet.ConditionalFormatting)
			{
				conditionalFormat.Address = new ExcelAddress(5 + i++, 5, 6, 6);
			}

			excelPackage.SaveAs(new FileInfo(@"c:\temp\error.xlsx"));
		}
	}
	[TestMethod, Ignore]
	public void Issue15436()
	{
		FileInfo file = new(@"c:\temp\incorrect value.xlsx");
		using ExcelPackage excelPackage = new(file);
		Assert.AreEqual(19120072, excelPackage.Workbook.Worksheets[1].Cells["A1"].Value);
	}
	[TestMethod, Ignore]
	public void Issue13128()
	{
		FileInfo file = new(@"c:\temp\students.xlsx");
		using ExcelPackage excelPackage = new(file);
		Assert.AreNotEqual(null, ((ExcelChart)excelPackage.Workbook.Worksheets[1].Drawings[0]).Series[0].XSeries);
	}
}
