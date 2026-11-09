using System.Text;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	internal ExcelColumn CopyColumn(ExcelColumn c, int col, int maxCol)
	{
		ExcelColumn newC = new(this, col)
		{
			ColumnMax = maxCol < ExcelPackage.MaxColumns ? maxCol : ExcelPackage.MaxColumns
		};
		if (c.StyleName != "")
		{
			newC.StyleName = c.StyleName;
		}
		else
		{
			newC.StyleID = c.StyleID;
		}

		newC.OutlineLevel = c.OutlineLevel;
		newC.Phonetic = c.Phonetic;
		newC.BestFit = c.BestFit;
		SetValueInner(0, col, newC);
		newC._width = c._width;
		newC._hidden = c._hidden;
		return newC;
	}
	/// <summary>
	/// Make the current worksheet active.
	/// </summary>
	public void Select() => View.TabSelected = true;
	/// <summary>
	/// Selects a range in the worksheet. The active cell is the topmost cell.
	/// Make the current worksheet active.
	/// </summary>
	/// <param name="Address">An address range</param>
	public void Select(string Address) => Select(Address, true);
	/// <summary>
	/// Selects a range in the worksheet. The actice cell is the topmost cell.
	/// </summary>
	/// <param name="Address">A range of cells</param>
	/// <param name="SelectSheet">Make the sheet active</param>
	public void Select(string Address, bool SelectSheet)
	{
		CheckSheetType();
		//Get rows and columns and validate as well
		ExcelCellBase.GetRowColFromAddress(Address, out var fromRow, out var fromCol, out int toRow, out var toCol);

		if (SelectSheet)
		{
			View.TabSelected = true;
		}

		View.SelectedRange = Address;
		View.ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);
	}
	/// <summary>
	/// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
	/// Make the current worksheet active.
	/// </summary>
	/// <param name="Address">An address range</param>
	public void Select(ExcelAddress Address)
	{
		CheckSheetType();
		Select(Address, true);
	}
	/// <summary>
	/// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
	/// </summary>
	/// <param name="Address">A range of cells</param>
	/// <param name="SelectSheet">Make the sheet active</param>
	public void Select(ExcelAddress Address, bool SelectSheet)
	{

		CheckSheetType();
		if (SelectSheet)
		{
			View.TabSelected = true;
		}

		var selAddress = new StringBuilder(ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column) + ":" + ExcelCellBase.GetAddress(Address.End.Row, Address.End.Column));
		if (Address.Addresses != null)
		{
			foreach (var a in Address.Addresses)
			{
				selAddress.Append(' ').Append(ExcelCellBase.GetAddress(a.Start.Row, a.Start.Column)).Append(':').Append(ExcelCellBase.GetAddress(a.End.Row, a.End.Column));
			}
		}

		View.SelectedRange = selAddress.ToString();
		View.ActiveCell = ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column);
	}

	/// <summary>
	/// Get the ExcelColumn for column (span ColumnMin and ColumnMax)
	/// </summary>
	/// <param name="column"></param>
	/// <returns></returns>
	internal ExcelColumn GetColumn(int column)
	{
		var c = GetValueInner(0, column) as ExcelColumn;
		if (c == null)
		{
			int row = 0, col = column;
			if (_values.PrevCell(ref row, ref col))
			{
				c = GetValueInner(0, col) as ExcelColumn;
				return c != null && c.ColumnMax >= column ? c : null;
			}
		}

		return c;

	}
}
