using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private static bool ParseQuote(string fullAddress, int i, bool isText, ref string first, ref string second, bool hasSheet)
	{
		if (isText && i + 1 < fullAddress.Length && fullAddress[i] == '\'')
		{
			AppendAddressChar(ref first, ref second, hasSheet, fullAddress[i]);
		}

		return !isText;
	}

	private bool ParseCharInBrackets(string fullAddress, int i, bool isText, string first, Stack<int> brackPos, List<string> bracketParts)
	{
		var c = fullAddress[i];
		if (c == '[' && !isText)
		{
			brackPos.Push(i);
		}
		else if (c == ']' && !isText)
		{
			if (brackPos.Count > 0)
			{
				var from = brackPos.Pop();
				bracketParts.Add(fullAddress.Substring(from + 1, i - from - 1));

				if (brackPos.Count == 0)
				{
					HandleBrackets(first, bracketParts);
				}
			}
			else
			{
				return false;
			}
		}

		return true;
	}

	private bool ParseCharOutsideBrackets(char c, int i, bool isText, ref string first, ref string second, ref bool hasSheet, bool isMulti, Stack<int> brackPos)
	{
		if (c == '[' && !isText)
		{
			brackPos.Push(i);
		}
		else if (IsSheetSeparator(c, isText, first, second))
		{
			second = RemoveRepeatedSheetName(first, second, hasSheet);
			hasSheet = true;
		}
		else if (c == ',' && !isText)
		{
			SetAddress(ref first, ref second, ref hasSheet, true);
			return true;
		}
		else
		{
			AppendAddressChar(ref first, ref second, hasSheet, c);
		}

		return isMulti;
	}

	// Sheet1!A1:Sheet1:A3 names the same worksheet twice and becomes Sheet1!A1:A3.
	private static string RemoveRepeatedSheetName(string first, string second, bool hasSheet)
	{
		if (hasSheet && second != null && second.ToLower(CultureInfo.CurrentCulture).EndsWith(first.ToLower(CultureInfo.CurrentCulture)))
		{
			return Regex.Replace(second, $"{first}$", string.Empty);
		}

		return second;
	}
	private static bool IsSheetSeparator(char c, bool isText, string first, string second)
		=> c == '!' && !isText && !first.EndsWith("#REF") && !second.EndsWith("#REF");

	private static void AppendAddressChar(ref string first, ref string second, bool hasSheet, char c)
	{
		if (hasSheet)
		{
			second += c;
		}
		else
		{
			first += c;
		}
	}
}
