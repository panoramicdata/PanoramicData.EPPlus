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
