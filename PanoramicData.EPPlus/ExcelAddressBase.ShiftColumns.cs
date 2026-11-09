namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	internal ExcelAddressBase AddColumn(int col, int cols, bool setFixed = false)
	{
		if (col > _toCol)
		{
			return this;
		}
		else
		{
			return col <= _fromCol
				? new ExcelAddressBase(_fromRow, ShiftIndex(_fromCol, _fromColFixed, setFixed, cols), _toRow, ShiftIndex(_toCol, _toColFixed, setFixed, cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed)
				: new ExcelAddressBase(_fromRow, _fromCol, _toRow, ShiftIndex(_toCol, _toColFixed, setFixed, cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
	}
	internal ExcelAddressBase DeleteColumn(int col, int cols, bool setFixed = false)
	{
		if (col > _toCol) //After
		{
			return this;
		}
		else if (col + cols <= _fromCol) //Before
		{
			return new ExcelAddressBase(_fromRow, ShiftIndex(_fromCol, _fromColFixed, setFixed, -cols), _toRow, ShiftIndex(_toCol, _toColFixed, setFixed, -cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		else if (col <= _fromCol && col + cols > _toCol) //Inside
		{
			return null;
		}
		else  //Partly
		{
			return col <= _fromCol
				? new ExcelAddressBase(_fromRow, col, _toRow, ShiftIndex(_toCol, _toColFixed, setFixed, -cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed)
				: new ExcelAddressBase(_fromRow, _fromCol, _toRow, ShrinkEnd(_toCol, _toColFixed, setFixed, cols, col), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
	}
}
