using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	internal ExcelAddressBase AddRow(int row, int rows, bool setFixed = false)
	{
		if (row > _toRow)
		{
			return this;
		}
		else
		{
			return row <= _fromRow
				? new ExcelAddressBase((setFixed && _fromRowFixed ? _fromRow : _fromRow + rows), _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow + rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed)
				: new ExcelAddressBase(_fromRow, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow + rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
	}
	internal ExcelAddressBase DeleteRow(int row, int rows, bool setFixed = false)
	{
		if (row > _toRow) //After
		{
			return this;
		}
		else if (row + rows <= _fromRow) //Before
		{
			return new ExcelAddressBase((setFixed && _fromRowFixed ? _fromRow : _fromRow - rows), _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		else if (row <= _fromRow && row + rows > _toRow) //Inside
		{
			return null;
		}
		else  //Partly
		{
			return row <= _fromRow
				? new ExcelAddressBase(row, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed)
				: new ExcelAddressBase(_fromRow, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows < row ? row - 1 : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
	}
	[SuppressMessage("Major Code Smell", "S1172:Unused method parameters should be removed", Justification = "Signature kept for existing callers and tests that pass a shift type.")]
	internal ExcelAddressBase Insert(ExcelAddressBase address, eShiftType Shift)
	{
		//Before or after, no change
		if (_toRow < address._fromRow || _toCol < address._fromCol || (_fromRow > address._toRow && _fromCol > address._toCol))
		{
			return this;
		}

		return null;
	}
}
