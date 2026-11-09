using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour of LoadFromText: callers may catch this exact exception type, so it is kept.")]
	private static int ParseTextLine(string line, ExcelTextFormat Format, List<object> items)
	{
		var col = 0;
		var v = new StringBuilder();
		bool isText = false, isQualifier = false;
		var QCount = 0;
		var lineQCount = 0;
		foreach (var c in line)
		{
			if (Format.TextQualifier != 0 && c == Format.TextQualifier)
			{
				if (!isText && v.Length != 0)
				{
					throw new Exception(string.Format("Invalid Text Qualifier in line : {0}", line));
				}

				isQualifier = !isQualifier;
				QCount += 1;
				lineQCount++;
				isText = true;
			}
			else
			{
				AppendPendingQualifiers(v, Format.TextQualifier, QCount);

				if (isQualifier)
				{
					v.Append(c);
				}
				else
				{
					AppendUnquotedChar(c, line, Format, v, items, QCount, ref isText, ref col);
				}

				QCount = 0;
			}
		}

		FinishTextLine(line, Format, v, items, QCount, lineQCount, isText, col);
		return col;
	}

	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour of LoadFromText: callers may catch this exact exception type, so it is kept.")]
	private static void AppendUnquotedChar(char c, string line, ExcelTextFormat Format, StringBuilder v, List<object> items, int QCount, ref bool isText, ref int col)
	{
		if (c == Format.Delimiter)
		{
			items.Add(ConvertData(Format, v.ToString(), col, isText));
			v.Clear();
			isText = false;
			col++;
		}
		else
		{
			if (QCount % 2 != 0)
			{
				throw new Exception(string.Format("Text delimiter is not closed in line : {0}", line));
			}

			v.Append(c);
		}
	}

	private static void AppendPendingQualifiers(StringBuilder v, char textQualifier, int QCount)
	{
		if (QCount > 1 && v.Length != 0)
		{
			v.Append(textQualifier, QCount / 2);
		}
		else if (QCount > 2 && v.Length == 0)
		{
			v.Append(textQualifier, (QCount - 1) / 2);
		}
	}

	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour of LoadFromText: callers may catch this exact exception type, so it is kept.")]
	private static void FinishTextLine(string line, ExcelTextFormat Format, StringBuilder v, List<object> items, int QCount, int lineQCount, bool isText, int col)
	{
		if (QCount > 1 && v.Length != 0 && QCount == 2)
		{
			v.Append(Format.TextQualifier, QCount / 2);
		}

		if (lineQCount % 2 != 0)
		{
			throw new Exception(string.Format("Text delimiter is not closed in line : {0}", line));
		}

		items.Add(ConvertData(Format, v.ToString(), col, isText));
	}
}
