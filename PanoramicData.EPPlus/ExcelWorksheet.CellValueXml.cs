using OfficeOpenXml.Compatibility;
using OfficeOpenXml.FormulaParsing;
using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private static string GetCellType(object v, bool allowStr = false)
	{
		if (v is bool)
		{
			return " t=\"b\"";
		}
		else if ((v is double v1 && double.IsInfinity(v1)) || v is ExcelErrorValue)
		{
			return " t=\"e\"";
		}
		else
		{
			return allowStr && v != null && !(TypeCompat.IsPrimitive(v) || v is double || v is decimal || v is DateTime || v is TimeSpan)
				? " t=\"str\""
				: "";
		}
	}

	private string GetValueForXml(object v)
	{
		string s;
		try
		{
			if (v is DateTime time)
			{
				var sdv = time.ToOADate();

				if (Workbook.Date1904)
				{
					sdv -= ExcelWorkbook.date1904Offset;
				}

				s = sdv.ToString(CultureInfo.InvariantCulture);
			}
			else if (v is TimeSpan span)
			{
				s = DateTime.FromOADate(0).Add(span).ToOADate().ToString(CultureInfo.InvariantCulture);
			}
			else if (TypeCompat.IsPrimitive(v) || v is double || v is decimal)
			{
				if (v is double v1 && double.IsNaN(v1))
				{
					s = "";
				}
				else if (v is double v2 && double.IsInfinity(v2))
				{
					s = "#NUM!";
				}
				else
				{
					s = Convert.ToDouble(v, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture);
				}
			}
			else
			{
				s = v.ToString();
			}
		}

		catch
		{
			s = "0";
		}

		return s;
	}
}
