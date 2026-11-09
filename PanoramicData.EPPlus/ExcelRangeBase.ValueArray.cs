namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private object GetValueArray()
	{
		ExcelAddressBase addr;
		if (_fromRow == 1 && _fromCol == 1 && _toRow == ExcelPackage.MaxRows && _toCol == ExcelPackage.MaxColumns)
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
				object o = null;
				if (_worksheet.ExistsValueInner(row, col, ref o))
				{
					if (_worksheet._flags.GetFlagValue(row, col, CellFlags.RichText))
					{
						v[row - addr._fromRow, col - addr._fromCol] = GetRichText(row, col).Text;
					}
					else
					{
						v[row - addr._fromRow, col - addr._fromCol] = o;
					}
				}
			}
		}

		return v;
	}
	private ExcelAddressBase GetAddressDim(ExcelRangeBase addr)
	{
		int fromRow, fromCol, toRow, toCol;
		var d = _worksheet.Dimension;
		fromRow = addr._fromRow < d._fromRow ? d._fromRow : addr._fromRow;
		fromCol = addr._fromCol < d._fromCol ? d._fromCol : addr._fromCol;

		toRow = addr._toRow > d._toRow ? d._toRow : addr._toRow;
		toCol = addr._toCol > d._toCol ? d._toCol : addr._toCol;

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
