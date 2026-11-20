using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test;

[TestClass]
public class Encrypt : TestBase
{
	[TestMethod]
	[Ignore]
	public void ReadWriteEncrypt()
	{
		using (ExcelPackage pck = new(new FileInfo(Path.Combine("Test", "Drawing.xlsx")), true))
		{
			pck.Encryption.Password = "EPPlus";
			pck.Encryption.Algorithm = EncryptionAlgorithm.AES192;
			pck.Workbook.Protection.SetPassword("test");
			pck.Workbook.Protection.LockStructure = true;
			pck.Workbook.Protection.LockWindows = true;

			pck.SaveAs(new FileInfo(Path.Combine("Test", "DrawingEncr.xlsx")));
		}

		using (ExcelPackage pck = new(new FileInfo(_worksheetPath + "DrawingEncr.xlsx"), true, "EPPlus"))
		{
			pck.Encryption.IsEncrypted = false;
			pck.SaveAs(new FileInfo(_worksheetPath + "DrawingNotEncr.xlsx"));
		}

		FileStream fs = new(_worksheetPath + "DrawingEncr.xlsx", FileMode.Open, FileAccess.ReadWrite);
		using (ExcelPackage pck = new(fs, "EPPlus"))
		{
			pck.Encryption.IsEncrypted = false;
			pck.SaveAs(new FileInfo(_worksheetPath + @"DrawingNotEncr.xlsx"));
		}

	}
	[TestMethod]
	[Ignore]
	public void WriteEncrypt()
	{
		var package = SharedTestData.CreateFirstLinePackage();

		package.Encryption.Password = "test";
		package.Encryption.IsEncrypted = true;
		package.SaveAs(new FileInfo(@"c:\temp\encrTest.xlsx"));
	}
	[TestMethod]
	[Ignore]
	public void WriteProtect()
	{
		ExcelPackage package = new(new FileInfo(@"c:\temp\workbookprot2.xlsx"), "");
		package.Workbook.Protection.SetPassword("t");
		package.Workbook.Protection.LockStructure = true;
		package.Workbook.View.Left = 585;
		package.Workbook.View.Top = 150;

		package.Workbook.View.Width = 17310;
		package.Workbook.View.Height = 38055;
		var ws = package.Workbook.Worksheets.Add("First line test");

		ws.Cells[1, 1].Value = "1; 1";
		ws.Cells[2, 1].Value = "2; 1";
		ws.Cells[1, 2].Value = "1; 2";
		ws.Cells[2, 2].Value = "2; 2";

		package.SaveAs(new FileInfo(@"c:\temp\workbookprot2.xlsx"));

	}
	[TestMethod]
	[Ignore]
	public void DecrypTest()
	{
		var p = new ExcelPackage(new FileInfo(@"c:\temp\encr.xlsx"), "test");

		_ = p.Workbook.Worksheets[1].Name;
		p.Encryption.Password = null;
		p.SaveAs(new FileInfo(@"c:\temp\encrNew.xlsx"));

	}
	[TestMethod, Ignore]
	public void DecrypTestBug()
	{
		var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\TestExcel_2040.xlsx"), "");

		_ = p.Workbook.Worksheets[1].Name;
		p.Encryption.Password = null;
		p.SaveAs(new FileInfo(@"c:\temp\encrNew.xlsx"));

	}
	[TestMethod]
	[Ignore]
	public void EncrypTest()
	{
		var f = new FileInfo(@"c:\temp\encrwrite.xlsx");
		if (f.Exists)
		{
			f.Delete();
		}

		var p = new ExcelPackage(f);

		p.Workbook.Protection.SetPassword("");
		p.Workbook.Protection.LockStructure = true;
		p.Encryption.Version = EncryptionVersion.Agile;

		var ws = p.Workbook.Worksheets.Add("Sheet1");
		for (var r = 1; r < 1000; r++)
		{
			ws.Cells[r, 1].Value = r;
		}

		p.Save();
	}
}
