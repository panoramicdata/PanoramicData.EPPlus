using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Utils;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Reads a row from the XML reader
	/// </summary>
	/// <param name="xr">The reader</param>
	/// <returns></returns>
	private static RowInternal AddRow(XmlReader xr) => new()
	{
		Collapsed = xr.GetAttribute("collapsed") is not null and "1",
		OutlineLevel = xr.GetAttribute("outlineLevel") == null ? (short)0 : short.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture),
		Height = xr.GetAttribute("ht") == null ? -1 : double.Parse(xr.GetAttribute("ht"), CultureInfo.InvariantCulture),
		Hidden = xr.GetAttribute("hidden") is not null and "1",
		Phonetic = xr.GetAttribute("ph") is not null and "1",
		CustomHeight = xr.GetAttribute("customHeight") is not null and "1"
	};

	private void SetValueFromXml(XmlReader xr, string type, int styleID, int row, int col)
	{
		if (type == "s")
		{
			var ix = xr.ReadElementContentAsInt();
			SetValueInner(row, col, _package.Workbook._sharedStringsList[ix].Text);
			if (_package.Workbook._sharedStringsList[ix].isRichText)
			{
				_flags.SetFlagValue(row, col, true, CellFlags.RichText);
			}
		}
		else if (type == "str")
		{
			SetValueInner(row, col, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
		}
		else if (type == "b")
		{
			SetValueInner(row, col, xr.ReadElementContentAsString() != "0");
		}
		else if (type == "e")
		{
			SetValueInner(row, col, GetErrorType(xr.ReadElementContentAsString()));
		}
		else
		{
			SetNumberValueFromXml(xr, styleID, row, col);
		}
	}

	private void SetNumberValueFromXml(XmlReader xr, int styleID, int row, int col)
	{
		var v = xr.ReadElementContentAsString();
		var nf = Workbook.Styles.CellXfs[styleID].NumberFormatId;
		if (nf is >= 14 and <= 22 or >= 45 and <= 47)
		{
			SetDateValueFromXml(v, row, col);
		}
		else
		{
			if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
			{
				SetValueInner(row, col, d);
			}
			else
			{
				SetValueInner(row, col, double.NaN);
			}
		}
	}

	private void SetDateValueFromXml(string v, int row, int col)
	{
		if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var res))
		{
			if (Workbook.Date1904)
			{
				res += ExcelWorkbook.date1904Offset;
			}

			if (res is >= -657435.0 and < 2958465.9999999)
			{
				SetValueInner(row, col, DateTime.FromOADate(res));
			}
			else
			{
				SetValueInner(row, col, res);
			}
		}
		else
		{
			SetValueInner(row, col, v);
		}
	}

	private static object GetErrorType(string v) => ExcelErrorValue.Parse(ConvertUtil._invariantTextInfo.ToUpper(v));
}
