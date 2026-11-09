using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private static object ConvertData(ExcelTextFormat Format, string v, int col, bool isText)
	{
		if (isText && (Format.DataTypes == null || Format.DataTypes.Length < col))
		{
			return NullIfEmpty(v);
		}

		if (Format.DataTypes == null || Format.DataTypes.Length <= col || Format.DataTypes[col] == eDataTypes.Unknown)
		{
			return ConvertUntypedData(Format, v);
		}
		else
		{
			return ConvertTypedData(Format.DataTypes[col], Format, v);
		}
	}

	private static string NullIfEmpty(string v) => string.IsNullOrEmpty(v) ? null : v;

	private static object ConvertUntypedData(ExcelTextFormat Format, string v)
	{
		var v2 = v.EndsWith('%') ? v[..^1] : v;
		if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out var d))
		{
			return v2 == v ? d : (object)(d / 100);
		}

		if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out var dt))
		{
			return dt;
		}
		else
		{
			return NullIfEmpty(v);
		}
	}

	private static object ConvertTypedData(eDataTypes dataType, ExcelTextFormat Format, string v)
	{
		switch (dataType)
		{
			case eDataTypes.Number:
				return double.TryParse(v, NumberStyles.Any, Format.Culture, out var d) ? d : v;
			case eDataTypes.DateTime:
				return DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out var dt) ? dt : v;
			case eDataTypes.Percent:
				return ConvertPercent(Format, v);
			case eDataTypes.String:
				return v;
			default:
				return NullIfEmpty(v);

		}
	}

	private static object ConvertPercent(ExcelTextFormat Format, string v)
	{
		var v2 = v.EndsWith('%') ? v[..^1] : v;
		return double.TryParse(v2, NumberStyles.Any, Format.Culture, out var d) ? d / 100 : v;
	}
}
