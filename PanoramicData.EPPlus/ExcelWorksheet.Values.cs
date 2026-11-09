using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Get the cell value from thw worksheet
	/// </summary>
	/// <param name="Row">The row number</param>
	/// <param name="Column">The row number</param>
	/// <returns>The value</returns>
	public object GetValue(int Row, int Column)
	{
		CheckSheetType();
		var v = GetValueInner(Row, Column);
		if (v != null)
		{
			return _flags.GetFlagValue(Row, Column, CellFlags.RichText) ? (object)Cells[Row, Column].RichText.Text : v;
		}
		else
		{
			return null;
		}
	}

	/// <summary>
	/// Get a strongly typed cell value from the worksheet
	/// </summary>
	/// <typeparam name="T">The type</typeparam>
	/// <param name="Row">The row number</param>
	/// <param name="Column">The row number</param>
	/// <returns>The value. If the value can't be converted to the specified type, the default value will be returned</returns>
	public T? GetValue<T>(int Row, int Column)
	{
		CheckSheetType();
		var v = GetValueInner(Row, Column);
		if (v == null)
		{
			return default;
		}

		return _flags.GetFlagValue(Row, Column, CellFlags.RichText)
			? (T)(object)Cells[Row, Column].RichText.Text
			: ConvertUtil.GetTypedCellValue<T>(v);
	}

	/// <summary>
	/// Set the value of a cell
	/// </summary>
	/// <param name="Row">The row number</param>
	/// <param name="Column">The column number</param>
	/// <param name="Value">The value</param>
	public void SetValue(int Row, int Column, object Value)
	{
		CheckSheetType();
		if (Row < 1 || Column < 1 || Row > ExcelPackage.MaxRows && Column > ExcelPackage.MaxColumns)
		{
			throw new ArgumentOutOfRangeException("Row or Column out of range");
		}

		SetValueInner(Row, Column, Value);
	}
	/// <summary>
	/// Set the value of a cell
	/// </summary>
	/// <param name="Address">The Excel address</param>
	/// <param name="Value">The value</param>
	public void SetValue(string Address, object Value)
	{
		CheckSheetType();
		ExcelCellBase.GetRowCol(Address, out var row, out var col, true);
		if (row < 1 || col < 1 || row > ExcelPackage.MaxRows && col > ExcelPackage.MaxColumns)
		{
			throw new ArgumentOutOfRangeException("Address is invalid or out of range");
		}

		SetValueInner(row, col, Value);
	}


	/// <summary>
	/// Get MergeCell Index No
	/// </summary>
	/// <param name="row"></param>
	/// <param name="column"></param>
	/// <returns></returns>
	public int GetMergeCellId(int row, int column)
	{
		for (var i = 0; i < _mergedCells.Count; i++)
		{
			if (!string.IsNullOrEmpty(_mergedCells[i]))
			{
				var range = Cells[_mergedCells[i]];

				if (range.Start.Row <= row && row <= range.End.Row)
				{
					if (range.Start.Column <= column && column <= range.End.Column)
					{
						return i + 1;
					}
				}
			}
		}

		return 0;
	}
}
