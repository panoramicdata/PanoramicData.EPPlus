using System.Text;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private static bool TrySplitAddress(string Address, out string wb, out string ws, out string intAddress)
	{
		wb = "";
		ws = "";
		intAddress = "";
		var text = new StringBuilder();
		var isText = false;
		var brackPos = -1;
		for (var i = 0; i < Address.Length; i++)
		{
			if (Address[i] == '\'')
			{
				isText = !isText;
				AppendEscapedQuote(Address, i, text);
			}
			else if (isText)
			{
				text.Append(Address[i]);
			}
			else if (Address[i] == '!')
			{
				SplitWorkbookAndSheet(text.ToString(), ref wb, out ws);
				intAddress = Address[(i + 1)..];
				return true;
			}
			else
			{
				var outcome = ParseBracketChar(Address, i, text, ref wb, ref brackPos);
				if (outcome.HasValue)
				{
					if (outcome.Value)
					{
						intAddress = Address;
					}

					return outcome.Value;
				}
			}
		}

		intAddress = text.ToString();
		return true;
	}

	private static void AppendEscapedQuote(string Address, int i, StringBuilder text)
	{
		if (i > 0 && Address[i - 1] == '\'')
		{
			text.Append('\'');
		}
	}

	private static void SplitWorkbookAndSheet(string s, ref string wb, out string ws)
	{
		if (s.Length > 0 && s[0] == '[')
		{
			wb = s[1..s.IndexOf(']')];
			ws = s[(s.IndexOf(']') + 1)..];
		}
		else
		{
			ws = s;
		}
	}

	// Null to keep reading; true for a table reference (the full address); false for an unmatched ']'.
	private static bool? ParseBracketChar(string Address, int i, StringBuilder text, ref string wb, ref int brackPos)
	{
		if (Address[i] == '[')
		{
			if (i > 0) // Table reference returns the full address
			{
				return true;
			}

			brackPos = i;
		}
		else if (Address[i] == ']')
		{
			if (brackPos > -1)
			{
				wb = text.ToString();
				text.Clear();
			}
			else
			{
				return false;
			}
		}
		else
		{
			text.Append(Address[i]);
		}

		return null;
	}
}
