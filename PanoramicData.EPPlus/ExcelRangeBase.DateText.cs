using OfficeOpenXml.Style.XmlAccess;
using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private static string GetDateText(DateTime d, string format, ExcelNumberFormatXml.ExcelFormatTranslator nf)
	{
		if (nf.SpecialDateFormat == ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemLongDate)
		{
			return d.ToLongDateString();
		}
		else if (nf.SpecialDateFormat == ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemLongTime)
		{
			return d.ToLongTimeString();
		}
		else if (nf.SpecialDateFormat == ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemShortDate)
		{
			return d.ToShortDateString();
		}

		return GetFormattedDateText(d, format, nf);
	}

	private static string GetFormattedDateText(DateTime d, string format, ExcelNumberFormatXml.ExcelFormatTranslator nf)
	{
		if (format is "d" or "D")
		{
			return d.Day.ToString();
		}
		else if (format == "M")
		{
			return d.Month.ToString();
		}
		else if (format == "m")
		{
			return d.Minute.ToString();
		}
		else if (format.ToLower(CultureInfo.CurrentCulture) is "y" or "yy")
		{
			return d.ToString("yy", nf.Culture);
		}
		else
		{
			return format.ToLower(CultureInfo.CurrentCulture) is "yyy" or "yyyy" ? d.ToString("yyy", nf.Culture) : d.ToString(format, nf.Culture);
		}
	}
}
