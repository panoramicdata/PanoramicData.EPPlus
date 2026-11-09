namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private static readonly _changeProp _setUnknownProp = SetUnknown;
	private static readonly _changeProp _setSingleProp = SetSingle;
	private static readonly _changeProp _setRangeProp = SetRange;
	private static readonly _changeProp _setMultiProp = SetMultiRange;
	private void SetDelegate()
	{
		if (_fromRow == -1)
		{
			_changePropMethod = SetUnknown;
		}
		//Single cell
		else if (_fromRow == _toRow && _fromCol == _toCol && Addresses == null)
		{
			_changePropMethod = SetSingle;
		}
		//Range (ex A1:A2)
		else if (Addresses == null)
		{
			_changePropMethod = SetRange;
		}
		//Multi Range (ex A1:A2,C1:C2)
		else
		{
			_changePropMethod = SetMultiRange;
		}
	}
	/// <summary>
	/// We dont know the address yet. Set the delegate first time a property is set.
	/// </summary>
	/// <param name="range"></param>
	/// <param name="valueMethod"></param>
	/// <param name="value"></param>
	private static void SetUnknown(ExcelRangeBase range, _setValue valueMethod, object? value)
	{
		//Address is not set use, selected range
		if (range._fromRow == -1)
		{
			range.SetToSelectedRange();
		}

		range.SetDelegate();
		range._changePropMethod(range, valueMethod, value);
	}
	/// <summary>
	/// Set a single cell
	/// </summary>
	/// <param name="range"></param>
	/// <param name="valueMethod"></param>
	/// <param name="value"></param>
	private static void SetSingle(ExcelRangeBase range, _setValue valueMethod, object? value) => valueMethod(range, value, range._fromRow, range._fromCol);
	/// <summary>
	/// Set a range
	/// </summary>
	/// <param name="range"></param>
	/// <param name="valueMethod"></param>
	/// <param name="value"></param>
	private static void SetRange(ExcelRangeBase range, _setValue valueMethod, object? value) => range.SetValueAddress(range, valueMethod, value);
	/// <summary>
	/// Set a multirange (A1:A2,C1:C2)
	/// </summary>
	/// <param name="range"></param>
	/// <param name="valueMethod"></param>
	/// <param name="value"></param>
	private static void SetMultiRange(ExcelRangeBase range, _setValue valueMethod, object? value)
	{
		range.SetValueAddress(range, valueMethod, value);
		foreach (var address in range.Addresses)
		{
			range.SetValueAddress(address, valueMethod, value);
		}
	}
	/// <summary>
	/// Set the property for an address
	/// </summary>
	/// <param name="address"></param>
	/// <param name="valueMethod"></param>
	/// <param name="value"></param>
	private void SetValueAddress(ExcelAddress address, _setValue valueMethod, object? value)
	{
		IsRangeValid("");
		if (IsWholeSheet())  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
		{
			throw new ArgumentException("Can't reference all cells. Please use the indexer to set the range");
		}
		else if (value is object[,] v && valueMethod == Set_Value)
		{
			// only simple set value is supported for bulk copy
			_worksheet.SetRangeValueInner(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, v);
		}
		else
		{
			SetValueForEachCell(address, valueMethod, value);
		}
	}

	private void SetValueForEachCell(ExcelAddress address, _setValue valueMethod, object? value)
	{
		for (var col = address.Start.Column; col <= address.End.Column; col++)
		{
			for (var row = address.Start.Row; row <= address.End.Row; row++)
			{
				valueMethod(this, value, row, col);
			}
		}
	}

	private bool IsWholeSheet()
		=> _fromRow == 1 && _fromCol == 1 && _toRow == ExcelPackage.MaxRows && _toCol == ExcelPackage.MaxColumns;
}
