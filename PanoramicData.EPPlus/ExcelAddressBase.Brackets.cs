using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private void HandleBrackets(string first, List<string> bracketParts)
	{
		if (!string.IsNullOrEmpty(first))
		{
			_table = new ExcelTableAddress();
			Table.Name = first;
			foreach (var s in bracketParts)
			{
				if (s.IndexOf('[') < 0)
				{
					ApplyTableSpecifier(s);
				}
			}
		}
	}

	private void ApplyTableSpecifier(string s)
	{
		switch (s.ToLower(CultureInfo.InvariantCulture))
		{
			case "#all":
				_table.IsAll = true;
				break;
			case "#headers":
				_table.IsHeader = true;
				break;
			case "#data":
				_table.IsData = true;
				break;
			case "#totals":
				_table.IsTotals = true;
				break;
			case "#this row":
				_table.IsThisRow = true;
				break;
			default:
				if (string.IsNullOrEmpty(_table.ColumnSpan))
				{
					_table.ColumnSpan = s;
				}
				else
				{
					_table.ColumnSpan += ":" + s;
				}

				break;
		}
	}
}
