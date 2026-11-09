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

		else
		{
			if (r1c1 && IsR1C1(Address))
			{
				return AddressType.R1C1;
			}
			else
			{
				if (SplitAddress(Address, out var wb, out var ws, out var intAddress))
				{
					if (intAddress.Contains('[')) //Table reference
					{
						return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
					}
					else if (intAddress.Contains(','))
					{
						intAddress = intAddress[..intAddress.IndexOf(',')];
					}

					return IsAddress(intAddress)
						? string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress
						: string.IsNullOrEmpty(wb) ? AddressType.InternalName : AddressType.ExternalName;
				}
				else
				{
					return AddressType.Invalid;
				}
			}
		}
	}

	private static bool IsR1C1(string address)
	{
		if (address.StartsWith('!'))
		{
			address = address[1..];
		}

		address = address.ToUpper();
		if (string.IsNullOrEmpty(address) || (address[0] != 'R' && address[0] != 'C'))
		{
			return false;
		}

		bool isC = false, isROrC = false;
		var startBracket = false;
		foreach (var c in address)
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
					if (c is >= '0' and <= '9' or '-')
					{
						if (isROrC == false)
						{
							return false;
						}
					}
					else
					{
						return false;
					}

					break;
			}
		}

		return true;
	}
}
