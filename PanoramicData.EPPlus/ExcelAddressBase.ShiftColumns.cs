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
				? new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed ? _fromCol : _fromCol + cols), _toRow, (setFixed && _toColFixed ? _toCol : _toCol + cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed)
				: new ExcelAddressBase(_fromRow, _fromCol, _toRow, (setFixed && _toColFixed ? _toCol : _toCol + cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
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
			return new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed ? _fromCol : _fromCol - cols), _toRow, (setFixed && _toColFixed ? _toCol : _toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		else if (col <= _fromCol && col + cols > _toCol) //Inside
		{
			return null;
		}
		else  //Partly
		{
			return col <= _fromCol
				? new ExcelAddressBase(_fromRow, col, _toRow, (setFixed && _toColFixed ? _toCol : _toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed)
				: new ExcelAddressBase(_fromRow, _fromCol, _toRow, (setFixed && _toColFixed ? _toCol : _toCol - cols < col ? col - 1 : _toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
	}
}
