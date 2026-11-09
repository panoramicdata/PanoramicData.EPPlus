using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

internal static class ThreeColumnSheet
{
	/// <summary>Adds sheet "test" with 1 to 4 in A1:A4 and 5 to 8 in both B1:B4 and C1:C4.</summary>
	public static ExcelWorksheet Add(ExcelPackage package)
	{
		var s1 = package.Workbook.Worksheets.Add("test");
		s1.Cells["A1"].Value = 1;
		s1.Cells["A2"].Value = 2;
		s1.Cells["A3"].Value = 3;
		s1.Cells["A4"].Value = 4;

		s1.Cells["B1"].Value = 5;
		s1.Cells["B2"].Value = 6;
		s1.Cells["B3"].Value = 7;
		s1.Cells["B4"].Value = 8;

		s1.Cells["C1"].Value = 5;
		s1.Cells["C2"].Value = 6;
		s1.Cells["C3"].Value = 7;
		s1.Cells["C4"].Value = 8;
		return s1;
	}
}