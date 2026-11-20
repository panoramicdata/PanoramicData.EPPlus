namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	internal void Delete(ExcelAddressBase Range, bool shift)
	{
		_worksheet.MergedCells.Clear(Range);
		//First find the start cell
		int fromRow, fromCol;
		var d = Worksheet.Dimension;
		if (d != null && Range._fromRow <= d._fromRow && Range._toRow >= d._toRow) //EntireRow?
		{
			fromRow = 0;
		}
		else
		{
			fromRow = Range._fromRow;
		}

		if (d != null && Range._fromCol <= d._fromCol && Range._toCol >= d._toCol) //EntireRow?
		{
			fromCol = 0;
		}
		else
		{
			fromCol = Range._fromCol;
		}

		var rows = Range._toRow - fromRow + 1;
		var cols = Range._toCol - fromCol + 1;

		_worksheet._values.Delete(fromRow, fromCol, rows, cols, shift);
		_worksheet._formulas.Delete(fromRow, fromCol, rows, cols, shift);
		_worksheet._hyperLinks.Delete(fromRow, fromCol, rows, cols, shift);
		_worksheet._flags.Delete(fromRow, fromCol, rows, cols, shift);
		_worksheet._commentsStore.Delete(fromRow, fromCol, rows, cols, shift);

		//Clear multi addresses as well
		if (Addresses != null)
		{
			foreach (var sub in Addresses)
			{
				Delete(sub, shift);
			}
		}
	}

	private void DeleteCheckMergedCells(ExcelAddressBase Range)
	{
		var removeItems = new List<string>();
		foreach (var addr in Worksheet.MergedCells)
		{
			var addrCol = Range.Collide(new ExcelAddress(Range.WorkSheet, addr));
			if (addrCol != eAddressCollition.No)
			{
				if (addrCol == eAddressCollition.Inside)
				{
					removeItems.Add(addr);
				}
				else
				{
					throw new InvalidOperationException("Can't remove/overwrite a part of cells that are merged");
				}
			}
		}

		foreach (var item in removeItems)
		{
			Worksheet.MergedCells.Remove(item);
		}
	}
}
