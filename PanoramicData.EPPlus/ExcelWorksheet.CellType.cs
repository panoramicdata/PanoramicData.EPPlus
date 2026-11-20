using OfficeOpenXml.Compatibility;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private static string GetCellType(object v, bool allowStr = false)
	{
		if (v is bool)
		{
			return " t=\"b\"";
		}
		else if (IsErrorCellValue(v))
		{
			return " t=\"e\"";
		}
		else
		{
			return allowStr && v != null && !IsNumericOrDateValue(v)
				? " t=\"str\""
				: "";
		}
	}

	private static bool IsErrorCellValue(object v) => (v is double v1 && double.IsInfinity(v1)) || v is ExcelErrorValue;

	private static bool IsNumericOrDateValue(object v) => TypeCompat.IsPrimitive(v) || v is double || v is decimal || v is DateTime || v is TimeSpan;
}
