using OfficeOpenXml;
namespace PanoramicData.EPPlus.Test;

public partial class ReadTemplate
{
	[Ignore]
	[TestMethod]
	public void StreamTest()
	{
		using var templateStream = File.OpenRead(@"c:\temp\thread.xlsx");

		using var outStream = File.Open(@"c:\temp\streamOut.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.None);
		using var package = new ExcelPackage(outStream, templateStream);
		package.Workbook.Worksheets[1].Cells["A1"].Value = 1;
		// Create more content
		package.Save();
	}
	[TestMethod, Ignore]
	public void test()
	{
		CreateXlsxSheet(@"C:\temp\bug\test4.xlsx", 4, 4);
		CreateXlsxSheet(@"C:\temp\bug\test25.xlsx", 25, 25);
	}
	[Ignore]
	[TestMethod]
	public void I15038()
	{
		using var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\15038.xlsx"));
		_ = p.Workbook.Worksheets[1];
	}
	[Ignore]
	[TestMethod]
	public void I15039()
	{
		using var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\15039.xlsm"));
		_ = p.Workbook.Worksheets[1];

		p.SaveAs(new FileInfo(@"c:\temp\bug\15039-saved.xlsm"));
	}
	[Ignore]
	[TestMethod]
	public void I15030()
	{
		using var newPack = new ExcelPackage(new FileInfo(@"c:\temp\bug\I15030.xlsx"));
		var wkBk = newPack.Workbook.Worksheets[1];
		var cell = wkBk.Cells["A1"];
		if (cell.Comment != null)
		{
			cell.Comment.Text = "Hello edited comments";
		}

		newPack.SaveAs(new FileInfo(@"c:\temp\bug\15030-save.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void I15014()
	{
		using var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\ClassicWineCompany.xlsx"));
		var ws = p.Workbook.Worksheets[1];
		Assert.AreEqual("SFFSectionHeading00", ws.Cells[5, 2].StyleName);
	}
	[Ignore]
	[TestMethod]
	public void I15043()
	{
		using var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\EPPlusTest\EPPlusTest\EPPlusTest\example.xlsx"));
		var ws = p.Workbook.Worksheets[1];
		p.Workbook.Worksheets.Copy(ws.Name, "Copy");
	}
	[TestMethod, Ignore]
	public void whitespace()
	{
		using var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\GridToExcel_05-12-2014.xlsx"));
		var ws = p.Workbook.Worksheets[1];
		foreach (var cell in ws.Cells[1, 84, 3, 86])
		{
			Console.WriteLine(cell.Address);
		}
	}
	[TestMethod, Ignore]
	public void SaveCorruption()
	{
		using var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\tables.xlsx"));
		var ws = p.Workbook.Worksheets[1];
		p.SaveAs(new FileInfo(@"c:\temp\bug\corr.xlsx"));
	}
	[TestMethod]
	public void VBAerror()
	{
		ExcelWorksheet ws;
		using var p = new ExcelPackage();
		p.Workbook.CreateVBAProject();
		ws = p.Workbook.Worksheets.Add("Градуировка");
		using var p2 = new ExcelPackage();
		p2.Workbook.CreateVBAProject();
		p2.Workbook.Worksheets.Add("Градуировка2", ws);
	}

	[TestMethod, Ignore]
	public void CopyIssue()
	{
		using var pkg = new ExcelPackage();
		var templateFile = ReadTemplateFile(@"C:\temp\bug\StackOverflow\EPPlusTest\20141120_01_3.各股累計收結表 (其他案件).xlsx");
		using var ms = new MemoryStream(templateFile);
		using var tempPkg = new ExcelPackage(ms);
		pkg.Workbook.Worksheets.Add("20141120_01_3.各股累計收結表 (其他案件)", tempPkg.Workbook.Worksheets.First());
	}
	[TestMethod, Ignore]
	public void FileStreamSave()
	{
		var fs = File.Create(@"c:\temp\fs.xlsx");
		using var pkg = new ExcelPackage(fs);
		var ws = pkg.Workbook.Worksheets.Add("test");
		ws.Cells["A1"].Value = 1;
		var col = ws.Column(1);
		col.OutlineLevel = 1;
		col.ColumnMax = ExcelPackage.MaxColumns;
		col.ColumnMax = 1;
		pkg.Save();
	}
}
