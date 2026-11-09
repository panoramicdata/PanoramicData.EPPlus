using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private void ExtractAddress(string fullAddress)
	{
		string first = "", second = "";
		bool hasSheet = false;
		try
		{
			if (fullAddress == "#REF!")
			{
				SetAddress(ref fullAddress, ref second, ref hasSheet, false);
				return;
			}
			else if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "!"))
			{
				// invalid address!
				return;
			}

			var isMulti = false;
			if (!ParseAddressCharacters(fullAddress, ref first, ref second, ref hasSheet, ref isMulti))
			{
				return;
			}

			if (Table == null)
			{
				SetAddress(ref first, ref second, ref hasSheet, isMulti);
			}
		}
		catch
		{
			// An unparsable address leaves the range unset
		}
	}

	private bool ParseAddressCharacters(string fullAddress, ref string first, ref string second, ref bool hasSheet, ref bool isMulti)
	{
		var brackPos = new Stack<int>();
		var bracketParts = new List<string>();
		var isText = false;
		for (var i = 0; i < fullAddress.Length; i++)
		{
			var c = fullAddress[i];
			if (c == '\'')
			{
				isText = ParseQuote(fullAddress, i, isText, ref first, ref second, hasSheet);
			}
			else if (brackPos.Count > 0)
			{
				if (!ParseCharInBrackets(fullAddress, i, isText, first, brackPos, bracketParts))
				{
					//Invalid address!
					return false;
				}
			}
			else
			{
				isMulti = ParseCharOutsideBrackets(c, i, isText, ref first, ref second, ref hasSheet, isMulti, brackPos);
			}
		}

		return true;
	}

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
			isMulti = true;
			SetAddress(ref first, ref second, ref hasSheet, isMulti);
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
			second = Regex.Replace(second, $"{first}$", string.Empty);
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

	private void SetAddress(ref string first, ref string second, ref bool hasSheet, bool isMulti)
	{
		string ws, address;
		if (hasSheet)
		{
			ws = first;
			address = second;
			first = "";
			second = "";
		}
		else
		{
			address = first;
			ws = "";
			first = "";
		}

		hasSheet = false;
		if (string.IsNullOrEmpty(_firstAddress))
		{
			if (string.IsNullOrEmpty(_ws) || !string.IsNullOrEmpty(ws))
			{
				_ws = ws;
			}

			_firstAddress = address;
			GetRowColFromAddress(address, out _fromRow, out _fromCol, out _toRow, out _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
		}

		if (isMulti)
		{
			_addresses ??= [];
			_addresses.Add(new ExcelAddress(_ws, address));
		}
		else
		{
			_addresses = null;
		}
	}
}
