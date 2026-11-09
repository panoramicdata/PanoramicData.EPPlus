namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	internal eAddressCollition Collide(ExcelAddressBase address, bool ignoreWs = false)
	{
		if (ignoreWs == false && address.WorkSheet != WorkSheet && address.WorkSheet != null)
		{
			return eAddressCollition.No;
		}

		if (address._fromRow > _toRow || address._fromCol > _toCol
			||
			_fromRow > address._toRow || _fromCol > address._toCol)
		{
			return eAddressCollition.No;
		}
		else if (address._fromRow == _fromRow && address._fromCol == _fromCol &&
				address._toRow == _toRow && address._toCol == _toCol)
		{
			return eAddressCollition.Equal;
		}
		else
		{
			return address._fromRow >= _fromRow && address._toRow <= _toRow &&
							 address._fromCol >= _fromCol && address._toCol <= _toCol
				? eAddressCollition.Inside
				: eAddressCollition.Partly;
		}
	}

	internal bool IsValidRowCol() => !(_fromRow > _toRow ||
		_fromCol > _toCol ||
		_fromRow < 1 ||
		_fromCol < 1 ||
		_toRow > ExcelPackage.MaxRows ||
		_toCol > ExcelPackage.MaxColumns);
}
