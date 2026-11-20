using OfficeOpenXml;
using System.Globalization;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[TestMethod]
	public void Issue345()
	{
		using var package = OpenTemplatePackage("issue345.xlsx");
		var worksheet = package.Workbook.Worksheets["test"];
		var sortColumns = new int[1];
		sortColumns[0] = 0;
		worksheet.Cells["A2:A30864"].Sort(sortColumns);
		package.Save();
	}
	[TestMethod]
	public void Issue387()
	{

		using var package = OpenTemplatePackage("issue345.xlsx");
		var workbook = package.Workbook;
		var worksheet = workbook.Worksheets.Add("One");

		worksheet.Cells[1, 3].Value = "Hello";
		var cells = worksheet.Cells["A3"];

		worksheet.Names.Add("R0", cells);
		workbook.Names.Add("Q0", cells);
	}
	[TestMethod]
	public void Issue333()
	{
		var previousCulture = CultureInfo.CurrentCulture;
		CultureInfo.CurrentCulture = new CultureInfo("sv-SE", false); // the built-in mm-dd-yy format is the culture's short date, which is ISO in sv-SE
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("TextBug");
		ws.Cells["A1"].Value = new DateTime(2019, 3, 7);
		ws.Cells["A1"].Style.Numberformat.Format = "mm-dd-yy";

		try
		{
			Assert.AreEqual("2019-03-07", ws.Cells["A1"].Text);
		}
		finally
		{
			CultureInfo.CurrentCulture = previousCulture;
		}
	}
	[TestMethod]
	public void Issue445()
	{
		ExcelPackage p = new();
		var ws = p.Workbook.Worksheets.Add("AutoFit"); //<-- This line takes forever. The process hangs.
		ws.Cells[1, 1].Value = new string('a', 50000);
		ws.Cells[1, 1].AutoFitColumns();
	}
	[TestMethod]
	public void Issue460()
	{
		var p = OpenTemplatePackage("Issue460.xlsx");
		var ws = p.Workbook.Worksheets[0];
		var newWs = p.Workbook.Worksheets.Add("NewSheet");
		ws.Cells.Copy(newWs.Cells);
		SaveWorksheet("Issue460_saved.xlsx");
	}
	[TestMethod]
	public void Issue476()
	{
		var p = OpenTemplatePackage("Issue345.xlsx");
		var ws = p.Workbook.Worksheets[0];
		var sortColumns = new int[1];
		sortColumns[0] = 0;
		ws.Cells["A2:A64515"].Sort(sortColumns);
		SaveWorksheet("Issue345_saved.xlsx");
	}
}
