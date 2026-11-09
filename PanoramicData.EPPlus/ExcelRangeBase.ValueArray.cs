namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private object GetValueArray()
	{
		ExcelAddressBase addr;
		if (IsWholeSheet())
		{
			addr = _worksheet.Dimension;
			if (addr == null)
			{
				return null;
			}
		}
		else
		{
			addr = this;
		}

		var v = new object[addr._toRow - addr._fromRow + 1, addr._toCol - addr._fromCol + 1];

		for (var col = addr._fromCol; col <= addr._toCol; col++)
		{
			for (var row = addr._fromRow; row <= addr._toRow; row++)
			{
				v[row - addr._fromRow, col - addr._fromCol] = GetArrayCellValue(row, col);
			}
		}

		return v;
	}

	private object GetArrayCellValue(int row, int col)
	{
		object o = null;
		if (!_worksheet.ExistsValueInner(row, col, ref o))
		{
			return null;
		}

		return _worksheet._flags.GetFlagValue(row, col, CellFlags.RichText) ? GetRichText(row, col).Text : o;
	}
	private ExcelAddressBase GetAddressDim(ExcelRangeBase addr)
	{
		int fromRow, fromCol, toRow, toCol;
		var d = _worksheet.Dimension;
		fromRow = Math.Max(addr._fromRow, d._fromRow);
		fromCol = Math.Max(addr._fromCol, d._fromCol);

		toRow = Math.Min(addr._toRow, d._toRow);
		toCol = Math.Min(addr._toCol, d._toCol);

		if (addr._fromRow == fromRow && addr._fromCol == fromCol && addr._toRow == toRow && addr._toCol == _toCol)
		{
			return addr;
		}
		else
		{
			return _fromRow > _toRow || _fromCol > _toCol ? null : new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
		}
	}

	private object GetSingleValue() => IsRichText ? RichText.Text : _worksheet.GetValueInner(_fromRow, _fromCol);
}
