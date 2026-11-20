using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

internal static class SheetFormulaHelper
{
	/// <summary>Sets the cells on a new sheet, calculates the formula and returns its value.</summary>
	internal static object CalculateOnNewSheet(string sheetName, string formulaCell, string formula, params (string Address, object Value)[] cells)
	{
		using var package = new ExcelPackage();
		var sheet = package.Workbook.Worksheets.Add(sheetName);
		foreach (var (address, value) in cells)
		{
			sheet.Cells[address].Value = value;
		}

		sheet.Cells[formulaCell].Formula = formula;
		sheet.Calculate();
		return sheet.Cells[formulaCell].Value;
	}
}