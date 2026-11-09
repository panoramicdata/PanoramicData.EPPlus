using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Style.XmlAccess;
using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private string GetFormattedText(bool forWidthCalc)
	{
		var v = Value;
		if (v == null)
		{
			return "";
		}

		var styles = Worksheet.Workbook.Styles;
		var nfID = styles.CellXfs[StyleID].NumberFormatId;
		ExcelNumberFormatXml.ExcelFormatTranslator nf = null;
		for (var i = 0; i < styles.NumberFormats.Count; i++)
		{
			if (nfID == styles.NumberFormats[i].NumFmtId)
			{
				nf = styles.NumberFormats[i].FormatTranslator;
				break;
			}
		}

		nf ??= styles.NumberFormats[0].FormatTranslator;  //nf should never be null. If so set to General, Issue 173

		string format, textFormat;
		if (forWidthCalc)
		{
			format = nf.NetFormatForWidth;
			textFormat = nf.NetTextFormatForWidth;
		}
		else
		{
			format = nf.NetFormat;
			textFormat = nf.NetTextFormat;
		}

		return FormatValue(v, nf, format, textFormat);
	}

	internal static string FormatValue(object v, ExcelNumberFormatXml.ExcelFormatTranslator nf, string format, string textFormat)
	{
		if (IsNumericValue(v))
		{
			return FormatNumericValue(v, nf, format);
		}
		else if (v is DateTime time)
		{
			return nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime
				? GetDateText(time, format, nf)
				: FormatNumber(time.ToOADate(), nf, format);
		}
		else if (v is TimeSpan span)
		{
			return nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime
				? GetDateText(new DateTime(span.Ticks), format, nf)
				: FormatNumber(new DateTime(0).Add(span).ToOADate(), nf, format);
		}
		else
		{
			return textFormat == "" ? v.ToString() : string.Format(textFormat, v);
		}
	}

	private static bool IsNumericValue(object v) => v is decimal || TypeCompat.IsPrimitive(v);

	private static string FormatNumericValue(object v, ExcelNumberFormatXml.ExcelFormatTranslator nf, string format)
	{
		double d;
		try
		{
			d = Convert.ToDouble(v);
		}
		catch
		{
			return "";
		}

		if (nf.DataType == ExcelNumberFormatXml.eFormatType.Number)
		{
			return FormatNumber(d, nf, format);
		}
		else if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
		{
			var date = DateTime.FromOADate(d);
			return GetDateText(date, format, nf);
		}

		return v.ToString();
	}

	private static string FormatNumber(double d, ExcelNumberFormatXml.ExcelFormatTranslator nf, string format)
		=> string.IsNullOrEmpty(nf.FractionFormat) ? d.ToString(format, nf.Culture) : nf.FormatFraction(d);

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
