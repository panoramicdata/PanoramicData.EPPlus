using OfficeOpenXml.Table;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	internal void SetRCFromTable(ExcelPackage pck, ExcelAddressBase? referenceAddress)
	{
		if (string.IsNullOrEmpty(_wb) && Table != null)
		{
			foreach (var ws in pck.Workbook.Worksheets)
			{
				foreach (var t in ws.Tables)
				{
					if (t.Name.Equals(Table.Name, StringComparison.OrdinalIgnoreCase))
					{
						_ws = ws.Name;
						SetRowsFromTable(t, referenceAddress);
						if (SetColumnsFromTable(t))
						{
							return;
						}
					}
				}
			}
		}
	}

	private bool SetColumnsFromTable(ExcelTable t)
	{
		if (string.IsNullOrEmpty(Table.ColumnSpan))
		{
			_fromCol = t.Address._fromCol;
			_toCol = t.Address._toCol;
			return true;
		}

		var col = t.Address._fromCol;
		var cols = Table.ColumnSpan.Split(':');
		foreach (var c in t.Columns)
		{
			if (_fromCol <= 0 && cols[0].Equals(c.Name, StringComparison.OrdinalIgnoreCase))   //Issue15063 Add invariant igore case
			{
				_fromCol = col;
				if (cols.Length == 1)
				{
					_toCol = _fromCol;
					return true;
				}
			}
			else if (IsColumnSpanEnd(cols, c.Name))
			{
				_toCol = col;
				return true;
			}

			col++;
		}

		return false;
	}

	private bool IsColumnSpanEnd(string[] cols, string name)
		=> cols.Length > 1 && _fromCol > 0 && cols[1].Equals(name, StringComparison.OrdinalIgnoreCase);
}
