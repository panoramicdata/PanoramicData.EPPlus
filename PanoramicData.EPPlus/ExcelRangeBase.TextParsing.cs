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

	private static string[] GetLines(string text, ExcelTextFormat Format)
	{
		if (Format.EOL == null || Format.EOL.Length == 0)
		{
			return [text];
		}

		var list = new List<string>();
		var prevLineStart = SplitOnEol(text, Format, list, out var inTQ);

		if (inTQ)
		{
			throw new ArgumentException(string.Format("Text delimiter is not closed in line : {0}", list.Count));
		}

		if (prevLineStart >= Format.EOL.Length && IsEOL(text, prevLineStart - Format.EOL.Length, Format.EOL))
		{
			list.Add("");
		}
		else
		{
			list.Add(text[prevLineStart..]);
		}

		return [.. list];
	}
	private static int SplitOnEol(string text, ExcelTextFormat Format, List<string> list, out bool inTQ)
	{
		var eol = Format.EOL;
		inTQ = false;
		var prevLineStart = 0;
		var i = 0;
		while (i < text.Length)
		{
			if (text[i] == Format.TextQualifier)
			{
				inTQ = !inTQ;
			}
			else if (!inTQ && IsEOL(text, i, eol))
			{
				list.Add(text[prevLineStart..i]);
				i += eol.Length - 1;
				prevLineStart = i + 1;
			}

			i++;
		}

		return prevLineStart;
	}

	private static bool IsEOL(string text, int ix, string eol)
	{
		for (var i = 0; i < eol.Length; i++)
		{
			if (text[ix + i] != eol[i])
			{
				return false;
			}
		}

		return ix + eol.Length <= text.Length;
	}
}
