using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

internal static class OneThreeSixSheet
{
	/// <summary>Adds sheet "Test" holding 1, 3 and 6 in A1:A3.</summary>
	public static ExcelWorksheet Add(ExcelPackage package)
	{
		var worksheet = package.Workbook.Worksheets.Add("Test");

		worksheet.Cells["A1"].Value = 1;
		worksheet.Cells["A2"].Value = 3;
		worksheet.Cells["A3"].Value = 6;
		return worksheet;
	}

	/// <summary>Sets the given values down a column from row 1.</summary>
	public static void SetColumn(ExcelWorksheet worksheet, string column, params object[] values)
	{
		for (var i = 0; i < values.Length; i++)
		{
			worksheet.Cells[column + (i + 1)].Value = values[i];
		}
	}
}