using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Get a range with an offset from the top left cell.
	/// The new range has the same dimensions as the current range
	/// </summary>
	/// <param name="RowOffset">Row Offset</param>
	/// <param name="ColumnOffset">Column Offset</param>
	/// <returns></returns>
	public ExcelRangeBase Offset(int RowOffset, int ColumnOffset)
	{
		if (_fromRow + RowOffset < 1 || _fromCol + ColumnOffset < 1 || _fromRow + RowOffset > ExcelPackage.MaxRows || _fromCol + ColumnOffset > ExcelPackage.MaxColumns)
		{
			throw new ArgumentOutOfRangeException("Offset value out of range");
		}

		var address = GetAddress(_fromRow + RowOffset, _fromCol + ColumnOffset, _toRow + RowOffset, _toCol + ColumnOffset);
		return new ExcelRangeBase(_worksheet, address);
	}
	/// <summary>
	/// Get a range with an offset from the top left cell.
	/// </summary>
	/// <param name="RowOffset">Row Offset</param>
	/// <param name="ColumnOffset">Column Offset</param>
	/// <param name="NumberOfRows">Number of rows. Minimum 1</param>
	/// <param name="NumberOfColumns">Number of columns. Minimum 1</param>
	/// <returns></returns>
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	public ExcelRangeBase Offset(int RowOffset, int ColumnOffset, int NumberOfRows, int NumberOfColumns)
	{
		if (NumberOfRows < 1 || NumberOfColumns < 1)
		{
			throw new Exception("Number of rows/columns must be greater than 0");
		}

		NumberOfRows--;
		NumberOfColumns--;
		if (_fromRow + RowOffset < 1 || _fromCol + ColumnOffset < 1 || _fromRow + RowOffset > ExcelPackage.MaxRows || _fromCol + ColumnOffset > ExcelPackage.MaxColumns ||
			 _fromRow + RowOffset + NumberOfRows < 1 || _fromCol + ColumnOffset + NumberOfColumns < 1 || _fromRow + RowOffset + NumberOfRows > ExcelPackage.MaxRows || _fromCol + ColumnOffset + NumberOfColumns > ExcelPackage.MaxColumns)
		{
			throw new ArgumentOutOfRangeException("Offset value out of range");
		}

		var address = GetAddress(_fromRow + RowOffset, _fromCol + ColumnOffset, _fromRow + RowOffset + NumberOfRows, _fromCol + ColumnOffset + NumberOfColumns);
		return new ExcelRangeBase(_worksheet, address);
	}
}
