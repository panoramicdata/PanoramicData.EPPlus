using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;
using System.Data;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[TestMethod, Ignore]
	public void Issue15252()
	{
		using var p = new ExcelPackage();
		var path1 = @"c:\temp\saveerror1.xlsx";
		var path2 = @"c:\temp\saveerror2.xlsx";
		var workSheet = p.Workbook.Worksheets.Add("saveerror");
		workSheet.Cells["A1"].Value = "test";

		// double save OK?
		p.SaveAs(new FileInfo(path1));
		p.SaveAs(new FileInfo(path2));

		// files are identical?
		var md5 = System.Security.Cryptography.SHA256.Create();
		using var fs1 = new FileStream(path1, FileMode.Open);
		using var fs2 = new FileStream(path2, FileMode.Open);
		var hash1 = string.Join("", md5.ComputeHash(fs1).Select((x) => { return x.ToString(); }));
		var hash2 = string.Join("", md5.ComputeHash(fs2).Select((x) => { return x.ToString(); }));
		Assert.AreEqual(hash1, hash2);
	}
	[TestMethod, Ignore]
	public void Issue15469()
	{
		ExcelPackage excelPackage = new(new FileInfo(@"c:\temp\bug\EPPlus-Bug.xlsx"), true);
		using FileStream fs = new(@"c:\temp\bug\EPPlus-Bug-new.xlsx", FileMode.Create);
		excelPackage.SaveAs(fs);
	}
	[TestMethod]
	public void Issue15438()
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("Test");
		var c = ws.Cells["A1"].Style.Font.Color;
		c.Indexed = 3;
		Assert.AreEqual("#FF00FF00", ExcelColor.LookupColor(c));
	}
	[TestMethod, Ignore]
	public void Issue15097()
	{
		using var pkg = new ExcelPackage();
		var templateFile = ReadTemplateFile(@"c:\temp\bug\test_vorlage3.xlsx");
		using var ms = new MemoryStream(templateFile);
		using var tempPkg = new ExcelPackage(ms);
		tempPkg.Workbook.Worksheets.Copy(tempPkg.Workbook.Worksheets.First().Name, "Demo");
	}
	[TestMethod, Ignore]
	public void Issue15485()
	{
		using var pkg = new ExcelPackage(new FileInfo(@"c:\temp\bug\PivotChartSeriesIssue.xlsx"));
		var ws = pkg.Workbook.Worksheets[1];
		ws.InsertRow(1, 1);
		ws.InsertColumn(1, 1);
		pkg.Save();
	}
	public static byte[] ReadTemplateFile(string templateName) => ReadTemplate.ReadTemplateFile(templateName);

	[TestMethod]
	public void Issue15455()
	{
		using var pck = new ExcelPackage();

		var sheet1 = pck.Workbook.Worksheets.Add("sheet1");
		var sheet2 = pck.Workbook.Worksheets.Add("Sheet2");
		sheet1.Cells["C2"].Value = 3;
		sheet1.Cells["C3"].Formula = "VLOOKUP(E1, Sheet2!A1:D6, C2, 0)";
		sheet1.Cells["E1"].Value = "d";

		sheet2.Cells["A1"].Value = "d";
		sheet2.Cells["C1"].Value = "dg";
		pck.Workbook.Calculate();
		var c3 = sheet1.Cells["C3"].Value;
		Assert.AreEqual("dg", c3);
	}

	private static object SaveAndReloadC3(object value)
	{
		FileInfo file = new("report.xlsx");
		try
		{
			if (file.Exists)
			{
				file.Delete();
			}
			using (ExcelPackage package = new(file))
			{
				var sheet = package.Workbook.Worksheets.Add("New Sheet");
				sheet.Cells[3, 3].Value = value;
				package.Save();
			}

			using (ExcelPackage package = new(file))
			{
				var sheet = package.Workbook.Worksheets["New Sheet"];
				return sheet.Cells[3, 3].Value;
			}
		}
		finally
		{
			if (file.Exists)
			{
				file.Delete();
			}
		}
	}

	[TestMethod]
	public void Issue15460WithString()
	{
		Assert.AreEqual("value1", SaveAndReloadC3(new[] { "value1", "value2", "value3" }));
	}

	[TestMethod]
	public void Issue15460WithNull()
	{
		Assert.AreEqual(string.Empty, SaveAndReloadC3(new[] { null, "value2", "value3" }));
	}

	[TestMethod]
	public void Issue15460WithNonStringPrimitive()
	{
		Assert.AreEqual((double)5, SaveAndReloadC3(new[] { 5, 6, 7 }));
	}
	[TestMethod, Ignore]
	public void MergeIssue()
	{
		var worksheetPath = Path.Combine(Path.GetTempPath(), @"EPPlus worksheets");
		FileInfo fi = new(Path.Combine(worksheetPath, "Example.xlsx"));
		fi.Delete();
		using ExcelPackage pckg = new(fi);
		var ws = pckg.Workbook.Worksheets.Add("Example");
		ws.Cells[1, 1, 1, 3].Merge = true;
		ws.Cells[1, 1, 1, 3].Merge = true;
		pckg.Save();
	}
}
