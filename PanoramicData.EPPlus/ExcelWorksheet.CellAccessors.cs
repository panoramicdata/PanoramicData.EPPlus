namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Get accessor of sheet value
	/// </summary>
	/// <param name="row">row</param>
	/// <param name="col">column</param>
	/// <returns>cell value</returns>
	internal object GetValueInner(int row, int col) => _values.GetValue(row, col)._value;
	/// <summary>
	/// Get accessor of sheet styleId
	/// </summary>
	/// <param name="row">row</param>
	/// <param name="col">column</param>
	/// <returns>cell styleId</returns>
	internal int GetStyleInner(int row, int col) => _values.GetValue(row, col)._styleId;

	/// <summary>
	/// Set accessor of sheet value
	/// </summary>
	/// <param name="row">row</param>
	/// <param name="col">column</param>
	/// <param name="value">value</param>
	internal void SetValueInner(int row, int col, object? value) => _values.SetValueSpecial(row, col, _setValueInnerUpdateDelegate, value);
	private static readonly CellStore<ExcelCoreValue>.SetValueDelegate _setValueInnerUpdateDelegate = SetValueInnerUpdate;
	private static void SetValueInnerUpdate(List<ExcelCoreValue> list, int index, object value) => list[index] = new ExcelCoreValue { _value = value, _styleId = list[index]._styleId };
	/// <summary>
	/// Set accessor of sheet styleId
	/// </summary>
	/// <param name="row">row</param>
	/// <param name="col">column</param>
	/// <param name="styleId">styleId</param>
	internal void SetStyleInner(int row, int col, int styleId) => _values.SetValueSpecial(row, col, (CellStore<ExcelCoreValue>.SetValueDelegate)SetStyleInnerUpdate, styleId);
	void SetStyleInnerUpdate(List<ExcelCoreValue> list, int index, object styleId) => list[index] = new ExcelCoreValue { _value = list[index]._value, _styleId = (int)styleId };

	/// <summary>
	/// Bulk(Range) set accessor of sheet value, for value array
	/// </summary>
	/// <param name="fromRow">start row</param>
	/// <param name="fromColumn">start column</param>
	/// <param name="toRow">end row</param>
	/// <param name="toColumn">end column</param>
	/// <param name="values">set values</param>
	internal void SetRangeValueInner(int fromRow, int fromColumn, int toRow, int toColumn, object[,] values)
	{
		var rowBound = values.GetUpperBound(0);
		var colBound = values.GetUpperBound(1);
		_values.SetRangeValueSpecial(fromRow, fromColumn, toRow, toColumn,
			(List<ExcelCoreValue> list, int index, int row, int column, object _) =>
			{
				object val = null;
				if (rowBound >= row - fromRow && colBound >= column - fromColumn)
				{
					val = values[row - fromRow, column - fromColumn];
				}

				list[index] = new ExcelCoreValue { _value = val, _styleId = list[index]._styleId };
			},
			values);
	}
	private static void SetRangeValueUpdate(List<ExcelCoreValue> list, int index, int row, int column, object values) => list[index] = new ExcelCoreValue { _value = ((object[,])values)[row, column], _styleId = list[index]._styleId };

	/// <summary>
	/// Existance check of sheet value
	/// </summary>
	/// <param name="row">row</param>
	/// <param name="col">column</param>
	/// <returns>is exists</returns>
	internal bool ExistsValueInner(int row, int col) => (_values.GetValue(row, col)._value != null);
	/// <summary>
	/// Existance check of sheet styleId
	/// </summary>
	/// <param name="row">row</param>
	/// <param name="col">column</param>
	/// <returns>is exists</returns>
	internal bool ExistsStyleInner(int row, int col) => (_values.GetValue(row, col)._styleId != 0);
	/// <summary>
	/// Existance check of sheet value
	/// </summary>
	/// <param name="row">row</param>
	/// <param name="col">column</param>
	/// <param name="value"></param>
	/// <returns>is exists</returns>
	internal bool ExistsValueInner(int row, int col, ref object value)
	{
		value = _values.GetValue(row, col)._value;
		return (value != null);
	}
	/// <summary>
	/// Existance check of sheet styleId
	/// </summary>
	/// <param name="row">row</param>
	/// <param name="col">column</param>
	/// <param name="styleId"></param>
	/// <returns>is exists</returns>
	internal bool ExistsStyleInner(int row, int col, ref int styleId)
	{
		styleId = _values.GetValue(row, col)._styleId;
		return (styleId != 0);
	}
}
