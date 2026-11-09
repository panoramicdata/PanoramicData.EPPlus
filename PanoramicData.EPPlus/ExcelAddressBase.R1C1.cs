using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private static bool IsR1C1(string address)
	{
		if (address.StartsWith('!'))
		{
			address = address[1..];
		}

		address = address.ToUpper(CultureInfo.CurrentCulture);
		if (string.IsNullOrEmpty(address) || (address[0] != 'R' && address[0] != 'C'))
		{
			return false;
		}

		bool isC = false, isROrC = false;
		var startBracket = false;
		foreach (var c in address)
		{
			if (!IsValidR1C1Char(c, ref isC, ref isROrC, ref startBracket))
			{
				return false;
			}
		}

		return true;
	}

	private static bool IsValidR1C1Char(char c, ref bool isC, ref bool isROrC, ref bool startBracket)
	{
		switch (c)
		{
			case 'C':
				isC = true;
				isROrC = true;
				break;
			case 'R':
				if (isC)
				{
					return false;
				}

				isROrC = true;
				break;
			case '[':
				startBracket = true;
				break;
			case ']':
				if (startBracket == false)
				{
					return false;
				}

				isROrC = false;
				break;
			case ':':
				isC = false;
				startBracket = false;
				isROrC = false;
				break;
			default:
				return IsR1C1Number(c, isROrC);
		}

		return true;
	}

	private static bool IsR1C1Number(char c, bool isROrC)
		=> (c is >= '0' and <= '9' or '-') && isROrC;
}
