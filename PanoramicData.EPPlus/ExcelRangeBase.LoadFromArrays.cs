namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Loads data from the collection of arrays of objects into the range, starting from
	/// the top-left cell.
	/// </summary>
	/// <param name="Data">The data.</param>
	public ExcelRangeBase LoadFromArrays(IEnumerable<object[]> Data)
	{
		//thanx to Abdullin for the code contribution
		ArgumentNullException.ThrowIfNull(Data);

		var rowArray = new List<object[]>();
		var maxColumn = 0;
		foreach (var item in Data)
		{
			rowArray.Add(item);
			if (maxColumn < item.Length)
			{
				maxColumn = item.Length;
			}
		}

		if (rowArray.Count == 0)
		{
			return null; //Issue #57
		}

		_worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + maxColumn - 1,
			(List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value) =>
			{
				rowIx -= _fromRow;
				columnIx -= _fromCol;

				var values = (List<object[]>)value;
				if (values.Count <= rowIx)
				{
					return;
				}

				var item = values[rowIx];
				if (item.Length <= columnIx)
				{
					return;
				}

				var val = item[columnIx];
				if (val != null && val != DBNull.Value && !string.IsNullOrEmpty(val.ToString()))
				{
					list[index] = new ExcelCoreValue { _value = val, _styleId = list[index]._styleId };
				}
			}, rowArray);

		return _worksheet.Cells[_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + maxColumn - 1];
	}
}
