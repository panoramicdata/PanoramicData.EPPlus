namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private Dictionary<int, int>? _columnStyles;

	// get StyleID without cell style for UpdateRowCellData
	internal int GetStyleIdDefaultWithMemo(int row, int col)
	{
		var v = 0;
		if (ExistsStyleInner(row, 0, ref v)) //First Row
		{
			return v;
		}
		else // then column
		{
			if (!_columnStyles.ContainsKey(col))
			{
				if (ExistsStyleInner(0, col, ref v))
				{
					_columnStyles.Add(col, v);
				}
				else
				{
					int r = 0, c = col;
					if (_values.PrevCell(ref r, ref c))
					{
						var val = _values.GetValue(0, c);
						var column = (ExcelColumn)(val._value);
						if (column != null && column.ColumnMax >= col) //Fixes issue 15174
						{
							_columnStyles.Add(col, val._styleId);
						}
						else
						{
							_columnStyles.Add(col, 0);
						}
					}
					else
					{
						_columnStyles.Add(col, 0);
					}
				}
			}

			return _columnStyles[col];
		}
	}
}
