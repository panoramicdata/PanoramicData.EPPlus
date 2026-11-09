using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Set the range to a specific value
	/// </summary>
	public object? Value
	{
		get
		{
			if (IsName)
			{
				return _worksheet == null ? _workbook._names[_address].NameValue : _worksheet.Names[_address].NameValue;
			}
			else
			{
				return _fromRow == _toRow && _fromCol == _toCol ? _worksheet.GetValue(_fromRow, _fromCol) : GetValueArray();
			}
		}
		set
		{
			if (IsName)
			{
				if (_worksheet == null)
				{
					_workbook._names[_address].NameValue = value;
				}
				else
				{
					_worksheet.Names[_address].NameValue = value;
				}
			}
			else
			{
				_changePropMethod(this, _setValueDelegate, value);
			}
		}
	}

	private static bool IsInfinityValue(object value)
	{
		var valueAsDouble = value as double?;

		return valueAsDouble.HasValue &&
			(double.IsNegativeInfinity(valueAsDouble.Value) || double.IsPositiveInfinity(valueAsDouble.Value));
	}

	/// <summary>
	/// Returns the formatted value.
	/// </summary>
	public string Text => GetFormattedText(false);
	/// <summary>
	/// Set the value without altering the richtext property
	/// </summary>
	/// <param name="value">the value</param>
	internal void SetValueRichText(object value)
	{
		if (_fromRow == 1 && _fromCol == 1 && _toRow == ExcelPackage.MaxRows && _toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
		{
			SetValue(value, 1, 1);
		}
		else
		{
			SetValue(value, _fromRow, _fromCol);
		}
	}

	private void SetValue(object value, int row, int col)
	{
		_worksheet.SetValue(row, col, value);
		_worksheet._formulas.SetValue(row, col, "");
	}

	/// <summary>
	///     Convert cell value to desired type, including nullable structs.
	///     When converting blank string to nullable struct (e.g. ' ' to int?) null is returned.
	///     When attempted conversion fails exception is passed through.
	/// </summary>
	/// <typeparam name="T">
	///     The type to convert to.
	/// </typeparam>
	/// <returns>
	///     The <see cref="Value"/> converted to <typeparamref name="T"/>.
	/// </returns>
	/// <remarks>
	///     If  <see cref="Value"/> is string, parsing is performed for output types of DateTime and TimeSpan, which if fails throws <see cref="FormatException"/>.
	///     Another special case for output types of DateTime and TimeSpan is when input is double, in which case <see cref="DateTime.FromOADate"/>
	///     is used for conversion. This special case does not work through other types convertible to double (e.g. integer or string with number).
	///     In all other cases 'direct' conversion <see cref="Convert.ChangeType(object, Type)"/> is performed.
	/// </remarks>
	/// <exception cref="FormatException">
	///      <see cref="Value"/> is string and its format is invalid for conversion (parsing fails)
	/// </exception>
	/// <exception cref="InvalidCastException">
	///      <see cref="Value"/> is not string and direct conversion fails
	/// </exception>
	public T GetValue<T>() => ConvertUtil.GetTypedCellValue<T>(Value);
}
