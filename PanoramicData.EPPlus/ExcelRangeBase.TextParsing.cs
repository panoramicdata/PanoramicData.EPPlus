namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
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
