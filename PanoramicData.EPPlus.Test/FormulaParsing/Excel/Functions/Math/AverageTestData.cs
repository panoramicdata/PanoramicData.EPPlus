using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Math;

/// <summary>The literal and cell-reference inputs shared by the AVERAGE and AVERAGEA tests.</summary>
internal static class AverageTestData
{
	public static double ExpectedLiteralAverage
	{
		get
		{
			double value1 = 1000;
			double value2 = 2000;
			double value3 = 6000;
			double value4 = 1;
			var value5 = new DateTime(2013, 1, 5).ToOADate();
			var value6 = new DateTime(2013, 1, 15).ToOADate();
			return (value1 + value2 + value3 + value4 + value5 + value6) / 6;
		}
	}

	public static CompileResult ExecuteOnLiterals(ExcelFunction average)
	{
		var date1 = new DateTime(2013, 1, 5);
		var date2 = new DateTime(2013, 1, 15);
		double value1 = 1000;
		double value2 = 2000;
		double value3 = 6000;
		return average.Execute(
		[
			new(value1.ToString("n")),
			new(value2),
			new(value3.ToString("n")),
			new(true),
			new(date1),
			new(date2.ToString("d"))
		], ParsingContext.Create());
	}

	public static CompileResult ExecuteOnCellReferences(ExcelFunction average)
	{
		ExcelPackage package = new();
		var worksheet = package.Workbook.Worksheets.Add("Test");
		var range1 = worksheet.Cells[1, 1];
		range1.Formula = "\"1000\"";
		range1.Calculate();
		var range2 = worksheet.Cells[1, 2];
		range2.Value = 2000;
		var range3 = worksheet.Cells[1, 3];
		range3.Formula = $"\"{new DateTime(2013, 1, 5):d}\"";
		range3.Calculate();
		var range4 = worksheet.Cells[1, 4];
		range4.Value = true;
		var range5 = worksheet.Cells[1, 5];
		range5.Value = new DateTime(2013, 1, 5);
		var range6 = worksheet.Cells[1, 6];
		range6.Value = "Test";
		var rangeInfo1 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 1, 1, 3);
		var rangeInfo2 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 4, 1, 4);
		var rangeInfo3 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 5, 1, 6);
		var context = ParsingContext.Create();
		var address = new OfficeOpenXml.FormulaParsing.ExcelUtilities.RangeAddress();
		address.FromRow = address.ToRow = address.FromCol = address.ToCol = 2;
		context.Scopes.NewScope(address);
		return average.Execute(
		[
			new(rangeInfo1),
			new(rangeInfo2),
			new(rangeInfo3)
		], context);
	}
}