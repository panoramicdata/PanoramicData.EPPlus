using OfficeOpenXml;
using OfficeOpenXml.Style.XmlAccess;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

/// <summary>Worksheet contents shared by several test classes.</summary>
internal static class SharedTestData
{
	/// <summary>Dates in U19:U24 and whole numbers in V19:V24.</summary>
	public static void AddDatesAndValues(ExcelWorksheet ws)
	{
		ws.Cells["U19"].Value = new DateTime(2009, 12, 31);
		ws.Cells["U20"].Value = new DateTime(2010, 1, 1);
		ws.Cells["U21"].Value = new DateTime(2010, 1, 2);
		ws.Cells["U22"].Value = new DateTime(2010, 1, 3);
		ws.Cells["U23"].Value = new DateTime(2010, 1, 4);
		ws.Cells["U24"].Value = new DateTime(2010, 1, 5);
		ws.Cells["U19:U24"].Style.Numberformat.Format = "yyyy-mm-dd";

		ws.Cells["V19"].Value = 100;
		ws.Cells["V20"].Value = 102;
		ws.Cells["V21"].Value = 101;
		ws.Cells["V22"].Value = 103;
		ws.Cells["V23"].Value = 105;
		ws.Cells["V24"].Value = 104;
	}

	/// <summary>A package whose sheet "First line test" has a bold first row and column over A1:B2.</summary>
	public static ExcelPackage CreateFirstLinePackage()
	{
		ExcelPackage package = new();
		var ws = package.Workbook.Worksheets.Add("First line test");

		ws.Cells[1, 1].Value = "1; 1";
		ws.Cells[2, 1].Value = "2; 1";
		ws.Cells[1, 2].Value = "1; 2";
		ws.Cells[2, 2].Value = "2; 2";

		ws.Row(1).Style.Font.Bold = true;
		ws.Column(1).Style.Font.Bold = true;
		return package;
	}

	/// <summary>Adds named style "Default" (Arial 18, underlined) and "Bold" cloned from it in red.</summary>
	public static (ExcelNamedStyleXml DefaultStyle, ExcelNamedStyleXml BoldStyle) AddArialDefaultAndRedBoldStyles(ExcelWorkbook workbook)
	{
		var defaultStyle = workbook.Styles.CreateNamedStyle("Default");
		defaultStyle.Style.Font.Name = "Arial";
		defaultStyle.Style.Font.Size = 18;
		defaultStyle.Style.Font.UnderLine = true;
		var boldStyle = workbook.Styles.CreateNamedStyle("Bold", defaultStyle.Style);
		boldStyle.Style.Font.Color.SetColor(Color.Red);
		return (defaultStyle, boldStyle);
	}
}
