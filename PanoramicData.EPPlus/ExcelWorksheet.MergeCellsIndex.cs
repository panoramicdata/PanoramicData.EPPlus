namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	public partial class MergeCellsCollection
	{
		private bool Validate(ExcelAddressBase address)
		{
			var ix = 0;
			if (_cells.Exists(address._fromRow, address._fromCol, ref ix))
			{
				return ix >= 0 && ix < _list.Count && _list[ix] != null && address.Address == _list[ix];
			}

			var cse = new CellsStoreEnumerator<int>(_cells, address._fromRow, address._fromCol, address._toRow, address._toCol);
			//cells
			while (cse.Next())
			{
				return false;
			}
			//Entire column
			cse = new CellsStoreEnumerator<int>(_cells, 0, address._fromCol, 0, address._toCol);
			while (cse.Next())
			{
				return false;
			}
			//Entire row
			cse = new CellsStoreEnumerator<int>(_cells, address._fromRow, 0, address._toRow, 0);
			while (cse.Next())
			{
				return false;
			}

			return true;
		}

		internal void SetIndex(ExcelAddressBase address, int ix)
		{
			int fromRow = address._fromRow, toRow = address._toRow, fromCol = address._fromCol, toCol = address._toCol;
			if (fromRow == 1 && toRow == ExcelPackage.MaxRows) //Entire column, stored on row 0
			{
				fromRow = toRow = 0;
			}
			else if (fromCol == 1 && toCol == ExcelPackage.MaxColumns) //Entire row, stored on column 0
			{
				fromCol = toCol = 0;
			}

			for (var col = fromCol; col <= toCol; col++)
			{
				for (var row = fromRow; row <= toRow; row++)
				{
					_cells.SetValue(row, col, ix);
				}
			}
		}
	}
}
