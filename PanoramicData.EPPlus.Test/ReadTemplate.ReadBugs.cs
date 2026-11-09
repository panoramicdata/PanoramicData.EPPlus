using System.Text;
using OfficeOpenXml;
namespace PanoramicData.EPPlus.Test;

public partial class ReadTemplate
{
	[Ignore]
	[TestMethod]
	public void ReadBug()
	{
		var file = new FileInfo(@"c:\temp\Adenoviridae Protocol.xlsx");
		using ExcelPackage pck = new(file);
		pck.Workbook.Worksheets[1].Cells["G4"].Value = 12;
		pck.SaveAs(new FileInfo(@"c:\temp\Adenoviridae Protocol2.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadBug3()
	{
		ExcelPackage xlsPack = new(new FileInfo(@"c:\temp\billing_template.xlsx"));
		var xlsWb = xlsPack.Workbook;
		_ = xlsWb.Worksheets["Billing"];
	}
	[Ignore]
	[TestMethod]
	public void ReadBug2()
	{
		var file = new FileInfo(@"c:\temp\book2.xlsx");
		using ExcelPackage pck = new(file);
		Assert.AreEqual("Good", pck.Workbook.Worksheets[1].Cells["A1"].StyleName);
		Assert.AreEqual("Good 2", pck.Workbook.Worksheets[1].Cells["C1"].StyleName);
		Assert.AreEqual("Note", pck.Workbook.Worksheets[1].Cells["G11"].StyleName);
		pck.SaveAs(new FileInfo(@"c:\temp\Adenoviridae Protocol2.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void CondFormatDataValBug()
	{
		var file = new FileInfo(@"c:\temp\condi.xlsx");
		using ExcelPackage pck = new(file);
		var dv = pck.Workbook.Worksheets[1].Cells["A1"].DataValidation.AddIntegerDataValidation();
		dv.Formula.Value = 1;
		dv.Formula2.Value = 4;
		dv.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.equal;
		pck.SaveAs(new FileInfo(@"c:\temp\condi2.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void InternalZip()
	{
		// Needs a local workbook that is not in the repository.
	}
	[Ignore]
	[TestMethod]
	public void ReadBug4()
	{
		var lines = new List<string>();
		var package = new ExcelPackage(new FileInfo(@"c:\temp\test.xlsx"));

		var workBook = package.Workbook;
		if (workBook != null)
		{
			if (workBook.Worksheets.Count > 0) //fails on this line
			{
				// Get the first worksheet
				var currentWorksheet = workBook.Worksheets.First();

				var rowCount = 1;
				var lastRow = currentWorksheet.Dimension.End.Row;
				var lastColumn = currentWorksheet.Dimension.End.Column;
				while (rowCount <= lastRow)
				{
					var columnCount = 1;
					var line = new StringBuilder();
					while (columnCount <= lastColumn)
					{
						line.Append(currentWorksheet.Cells[rowCount, columnCount].Value).Append('|');
						columnCount++;
					}

					lines.Add(line.ToString());
					rowCount++;
				}
			}
		}
	}
	[Ignore]
	[TestMethod]
	public void ReadBug5()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\protect.xlsx"));

		package.Workbook.Worksheets[1].Protection.AllowInsertColumns = true;
		package.Workbook.Worksheets[1].Protection.SetPassword("test");
		package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\protectnew.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadBug6()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\outofrange\error.xlsx"));

		package.Workbook.Worksheets[1].Protection.AllowInsertColumns = true;
		package.Workbook.Worksheets[1].Protection.SetPassword("test");
		package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\error.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadBug7()
	{
		var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("test");
		using (var rng = ws.Cells["A1"])
		{
			var rt1 = rng.RichText.Add("TEXT1\r\n");
			rt1.Bold = true;
			rng.Style.WrapText = true;
			var rt2 = rng.RichText.Add("TEXT2");
			rt2.Bold = false;
		}

		package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\error.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadBug8()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\bug\Genband SO CrossRef Phoenix.xlsx"));
		var ws = package.Workbook.Worksheets[1];
		using (var rng = ws.Cells["A1"])
		{
			var rt1 = rng.RichText.Add("TEXT1\r\n");
			rt1.Bold = true;
			rng.Style.WrapText = true;
			var rt2 = rng.RichText.Add("TEXT2");
			rt2.Bold = false;
		}

		package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\billing_template.xlsx.error"));
	}
}
