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
						if (Table.IsAll)
						{
							_fromRow = t.Address._fromRow;
							_toRow = t.Address._toRow;
						}
						else
						{
							if (Table.IsThisRow)
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
							else if (Table.IsHeader && Table.IsData)
							{
								_fromRow = t.Address._fromRow;
								_toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
							}
							else if (Table.IsData && Table.IsTotals)
							{
								_fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
								_toRow = t.Address._toRow;
							}
							else if (Table.IsHeader)
							{
								_fromRow = t.ShowHeader ? t.Address._fromRow : -1;
								_toRow = t.ShowHeader ? t.Address._fromRow : -1;
							}
							else if (Table.IsTotals)
							{
								_fromRow = t.ShowTotal ? t.Address._toRow : -1;
								_toRow = t.ShowTotal ? t.Address._toRow : -1;
							}
							else
							{
								_fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
								_toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
							}
						}

						if (string.IsNullOrEmpty(Table.ColumnSpan))
						{
							_fromCol = t.Address._fromCol;
							_toCol = t.Address._toCol;
							return;
						}
						else
						{
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
										return;
									}
								}
								else if (cols.Length > 1 && _fromCol > 0 && cols[1].Equals(c.Name, StringComparison.OrdinalIgnoreCase)) //Issue15063 Add invariant igore case
								{
									_toCol = col;
									return;
								}

								col++;
							}
						}
					}
				}
			}
		}
	}
}
