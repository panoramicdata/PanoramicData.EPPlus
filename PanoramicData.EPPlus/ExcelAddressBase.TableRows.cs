using OfficeOpenXml.Table;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private void SetRowsFromTable(ExcelTable t, ExcelAddressBase? referenceAddress)
	{
		if (Table.IsAll)
		{
			_fromRow = t.Address._fromRow;
			_toRow = t.Address._toRow;
		}
		else if (Table.IsThisRow)
		{
			SetRowsFromThisRow(referenceAddress);
		}
		else
		{
			SetRowsFromTableSection(t);
		}
	}

	private void SetRowsFromThisRow(ExcelAddressBase? referenceAddress)
	{
		if (referenceAddress == null)
		{
			_fromRow = -1;
			_toRow = -1;
		}
		else
		{
			_fromRow = referenceAddress._fromRow;
			_toRow = _fromRow;
		}
	}

	private void SetRowsFromTableSection(ExcelTable t)
	{
		if (Table.IsHeader && Table.IsData)
		{
			_fromRow = t.Address._fromRow;
			_toRow = LastDataRow(t);
		}
		else if (Table.IsData && Table.IsTotals)
		{
			_fromRow = FirstDataRow(t);
			_toRow = t.Address._toRow;
		}
		else if (Table.IsHeader)
		{
			_fromRow = HeaderRow(t);
			_toRow = HeaderRow(t);
		}
		else if (Table.IsTotals)
		{
			_fromRow = TotalsRow(t);
			_toRow = TotalsRow(t);
		}
		else
		{
			_fromRow = FirstDataRow(t);
			_toRow = LastDataRow(t);
		}
	}

	private static int FirstDataRow(ExcelTable t) => t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;

	private static int LastDataRow(ExcelTable t) => t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;

	private static int HeaderRow(ExcelTable t) => t.ShowHeader ? t.Address._fromRow : -1;

	private static int TotalsRow(ExcelTable t) => t.ShowTotal ? t.Address._toRow : -1;
}
