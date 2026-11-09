using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	internal enum AddressType
	{
		Invalid,
		InternalAddress,
		ExternalAddress,
		InternalName,
		ExternalName,
		Formula,
		R1C1
	}

	internal static AddressType IsValid(string Address, bool r1c1 = false)
	{
		if (Address == "#REF!")
		{
			return AddressType.Invalid;
		}
		else if (double.TryParse(Address, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) //A double, no valid address
		{
			return AddressType.Invalid;
		}
		else if (IsFormula(Address))
		{
			return AddressType.Formula;
		}
		else if (r1c1 && IsR1C1(Address))
		{
			return AddressType.R1C1;
		}
		else
		{
			return GetAddressOrNameType(Address);
		}
	}

	private static AddressType GetAddressOrNameType(string Address)
	{
		if (TrySplitAddress(Address, out var wb, out var ws, out var intAddress))
		{
			if (intAddress.Contains('[')) //Table reference
			{
				return GetAddressType(wb);
			}
			else if (intAddress.Contains(','))
			{
				intAddress = intAddress[..intAddress.IndexOf(',')];
			}

			return IsAddress(intAddress)
				? GetAddressType(wb)
				: GetNameType(wb);
		}
		else
		{
			return AddressType.Invalid;
		}
	}

	private static AddressType GetAddressType(string wb)
		=> string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;

	private static AddressType GetNameType(string wb)
		=> string.IsNullOrEmpty(wb) ? AddressType.InternalName : AddressType.ExternalName;

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
