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

	private void SetRowsFromTable(ExcelTable t, ExcelAddressBase? referenceAddress)
	{
		if (Table.IsAll)
		{
			_fromRow = t.Address._fromRow;
			_toRow = t.Address._toRow;
		}
		else if (Table.IsThisRow)
		{
			SetRowsFromThisRow(referenceAddress);
		}
		else
		{
			SetRowsFromTableSection(t);
		}
	}

	private void SetRowsFromThisRow(ExcelAddressBase? referenceAddress)
	{
		if (referenceAddress == null)
		{
			_fromRow = -1;
			_toRow = -1;
		}
		else
		{
			_fromRow = referenceAddress._fromRow;
			_toRow = _fromRow;
		}
	}

	private void SetRowsFromTableSection(ExcelTable t)
	{
		if (Table.IsHeader && Table.IsData)
		{
			_fromRow = t.Address._fromRow;
			_toRow = LastDataRow(t);
		}
		else if (Table.IsData && Table.IsTotals)
		{
			_fromRow = FirstDataRow(t);
			_toRow = t.Address._toRow;
		}
		else if (Table.IsHeader)
		{
			_fromRow = HeaderRow(t);
			_toRow = HeaderRow(t);
		}
		else if (Table.IsTotals)
		{
			_fromRow = TotalsRow(t);
			_toRow = TotalsRow(t);
		}
		else
		{
			_fromRow = FirstDataRow(t);
			_toRow = LastDataRow(t);
		}
	}

	private static int FirstDataRow(ExcelTable t) => t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;

	private static int LastDataRow(ExcelTable t) => t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;

	private static int HeaderRow(ExcelTable t) => t.ShowHeader ? t.Address._fromRow : -1;

	private static int TotalsRow(ExcelTable t) => t.ShowTotal ? t.Address._toRow : -1;

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
