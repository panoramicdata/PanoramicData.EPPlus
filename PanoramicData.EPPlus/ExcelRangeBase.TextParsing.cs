using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private static object ConvertData(ExcelTextFormat Format, string v, int col, bool isText)
	{
		if (isText && (Format.DataTypes == null || Format.DataTypes.Length < col))
		{
			return string.IsNullOrEmpty(v) ? null : v;
		}

		double d;
		DateTime dt;
		if (Format.DataTypes == null || Format.DataTypes.Length <= col || Format.DataTypes[col] == eDataTypes.Unknown)
		{
			var v2 = v.EndsWith('%') ? v[..^1] : v;
			if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out d))
			{
				return v2 == v ? d : (object)(d / 100);
			}

			if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out dt))
			{
				return dt;
			}
			else
			{
				return string.IsNullOrEmpty(v) ? null : v;
			}
		}
		else
		{
			switch (Format.DataTypes[col])
			{
				case eDataTypes.Number:
					return double.TryParse(v, NumberStyles.Any, Format.Culture, out d) ? d : v;
				case eDataTypes.DateTime:
					return DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out dt) ? dt : v;
				case eDataTypes.Percent:
					var v2 = v.EndsWith('%') ? v[..^1] : v;
					return double.TryParse(v2, NumberStyles.Any, Format.Culture, out d) ? d / 100 : v;
				case eDataTypes.String:
					return v;
				default:
					return string.IsNullOrEmpty(v) ? null : v;

			}
		}
	}

	private static string[] GetLines(string text, ExcelTextFormat Format)
	{
		if (Format.EOL == null || Format.EOL.Length == 0)
		{
			return [text];
		}

		var eol = Format.EOL;
		var list = new List<string>();
		var inTQ = false;
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
