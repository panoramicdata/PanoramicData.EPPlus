using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

/// <summary>Runs a criteria function such as SUMIF over A1:A3, optionally summing B1:B3 (1, 3, 5).</summary>
internal sealed class CriteriaRangeHelper(EpplusExcelDataProvider provider, ExcelWorksheet worksheet, ParsingContext parsingContext)
{
	public object Execute(ExcelFunction func, object? a1, object? a2, object? a3, object? criteria)
	{
		SetCriteriaColumn(a1, a2, a3);
		var range = provider.GetRange(worksheet.Name, 1, 1, 3, 1);
		var args = FunctionsHelper.CreateArgs(range, criteria);
		return func.Execute(args, parsingContext).Result;
	}

	public object ExecuteWithSumRange(ExcelFunction func, object? a1, object? a2, object? a3, object? criteria)
	{
		SetCriteriaColumn(a1, a2, a3);
		worksheet.Cells["B1"].Value = 1d;
		worksheet.Cells["B2"].Value = 3d;
		worksheet.Cells["B3"].Value = 5d;
		var range1 = provider.GetRange(worksheet.Name, 1, 1, 3, 1);
		var range2 = provider.GetRange(worksheet.Name, 1, 2, 3, 2);
		var args = FunctionsHelper.CreateArgs(range1, criteria, range2);
		return func.Execute(args, parsingContext).Result;
	}

	private void SetCriteriaColumn(object? a1, object? a2, object? a3)
	{
		worksheet.Cells["A1"].Value = a1;
		worksheet.Cells["A2"].Value = a2;
		worksheet.Cells["A3"].Value = a3;
	}
}
