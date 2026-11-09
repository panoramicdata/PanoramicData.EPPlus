using System.Text;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private static bool SplitAddress(string Address, out string wb, out string ws, out string intAddress)
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
				if (i > 0 && Address[i - 1] == '\'')
				{
					text.Append('\'');
				}
			}
			else
			{
				if (Address[i] == '!' && !isText)
				{
					var s = text.ToString();
					if (s.Length > 0 && s[0] == '[')
					{
						wb = s[1..s.IndexOf(']')];
						ws = s[(s.IndexOf(']') + 1)..];
					}
					else
					{
						ws = s;
					}

					intAddress = Address[(i + 1)..];
					return true;
				}
				else
				{
					if (Address[i] == '[' && !isText)
					{
						if (i > 0) // Table reference returns the full address
						{
							intAddress = Address;
							return true;
						}

						brackPos = i;
					}
					else if (Address[i] == ']' && !isText)
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
				}
			}
		}

		intAddress = text.ToString();
		return true;
	}
	internal static String GetWorkbookPart(string address)
	{
		var ix = 0;
		if (address[0] == '[')
		{
			ix = address.IndexOf(']') + 1;
			if (ix > 0)
			{
				return address.Substring(1, ix - 2);
			}
		}

		return "";
	}
	internal static string GetWorksheetPart(string address, string defaultWorkSheet)
	{
		var ix = 0;
		return GetWorksheetPart(address, defaultWorkSheet, ref ix);
	}
	internal static string GetWorksheetPart(string address, string defaultWorkSheet, ref int endIx)
	{
		if (address == "")
		{
			return defaultWorkSheet;
		}

		var ix = 0;
		if (address[0] == '[')
		{
			ix = address.IndexOf(']') + 1;
		}

		if (ix > 0 && ix < address.Length)
		{
			if (address[ix] == '\'')
			{
				return GetString(address, ix, out endIx);
			}
			else
			{
				var ixEnd = address.IndexOf('!', ix);
				return ixEnd > ix ? address[ix..ixEnd] : defaultWorkSheet;
			}
		}
		else
		{
			return defaultWorkSheet;
		}
	}
	internal static string GetAddressPart(string address)
	{
		var ix = 0;
		GetWorksheetPart(address, "", ref ix);
		if (ix < address.Length)
		{
			return address[ix] == '!' ? address[(ix + 1)..] : "";
		}
		else
		{
			return "";
		}
	}
	internal static void SplitAddress(string fullAddress, out string wb, out string ws, out string address, string defaultWorksheet = "")
	{
		wb = GetWorkbookPart(fullAddress);
		var ix = 0;
		ws = GetWorksheetPart(fullAddress, defaultWorksheet, ref ix);
		if (ix < fullAddress.Length)
		{
			if (fullAddress[ix] == '!')
			{
				address = fullAddress[(ix + 1)..];
			}
			else
			{
				address = fullAddress[ix..];
			}
		}
		else
		{
			address = "";
		}
	}
	private static string GetString(string address, int ix, out int endIx)
	{
		var strIx = address.IndexOf("''", StringComparison.CurrentCulture);
		while (strIx > -1)
		{
			strIx = address.IndexOf("''", StringComparison.CurrentCulture);
		}

		endIx = address.IndexOf('\'');
		return address[ix..endIx].Replace("''", "'");
	}
}
