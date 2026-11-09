using System.Text.RegularExpressions;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private static bool IsAddress(string intAddress)
	{
		if (string.IsNullOrEmpty(intAddress))
		{
			return false;
		}

		var cells = intAddress.Split(':');
		int toRow, toCol;

		if (!GetRowCol(cells[0], out var fromRow, out var fromCol, false))
		{
			return false;
		}

		if (cells.Length > 1)
		{
			if (!GetRowCol(cells[1], out toRow, out toCol, false))
			{
				return false;
			}
		}
		else
		{
			toRow = fromRow;
			toCol = fromCol;
		}

		return fromRow <= toRow &&
			fromCol <= toCol &&
			fromCol > -1 &&
			toCol <= ExcelPackage.MaxColumns &&
			fromRow > -1 &&
			toRow <= ExcelPackage.MaxRows;
	}

	private static readonly HashSet<char> _formulaCharacters = new(['(', ')', '+', '-', '*', '/', '=', '^', '&', '%', '\"']);

	private static bool IsFormula(string address)
	{
		var isText = false;
		for (var i = 0; i < address.Length; i++)
		{
			var addressChar = address[i];
			if (addressChar == '\'')
			{
				isText = !isText;
			}
			else
			{
				// Table references use [ ] around column names and since table column names can also contain unescaped formula characters,
				// we need to check that this is not a table column reference in order to avoid false positives.  Since function names and
				// formulas cannot contain [ ], we should be safe doing this check.
				if (addressChar is '[' or ']')
				{
					return false;
				}

				if (isText == false && _formulaCharacters.Contains(addressChar))
				{
					return true;
				}
			}
		}

		return false;
	}

	private static bool IsValidName(string address) => Regex.IsMatch(address, "[^0-9./*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|][^/*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|]*");
}
