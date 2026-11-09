using OfficeOpenXml.DataValidation;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Inserts a new row into the spreadsheet.  Existing rows below the position are 
	/// shifted down.  All formula are updated to take account of the new row.
	/// </summary>
	/// <param name="rowFrom">The position of the new row</param>
	/// <param name="rows">Number of rows to insert</param>
	public void InsertRow(int rowFrom, int rows) => InsertRow(rowFrom, rows, 0);
	/// <summary>
	/// Inserts a new row into the spreadsheet.  Existing rows below the position are 
	/// shifted down.  All formula are updated to take account of the new row.
	/// </summary>
	/// <param name="rowFrom">The position of the new row</param>
	/// <param name="rows">Number of rows to insert.</param>
	/// <param name="copyStylesFromRow">Copy Styles from this row. Applied to all inserted rows</param>
	public void InsertRow(int rowFrom, int rows, int copyStylesFromRow)
	{
		CheckSheetType();
		var d = Dimension;

		if (rowFrom < 1)
		{
			throw new ArgumentOutOfRangeException("rowFrom can't be lesser that 1");
		}

		//Check that cells aren't shifted outside the boundries
		if (d != null && d.End.Row > rowFrom && d.End.Row + rows > ExcelPackage.MaxRows)
		{
			throw new ArgumentOutOfRangeException("Can't insert. Rows will be shifted outside the boundries of the worksheet.");
		}

		lock (this)
		{
			_values.Insert(rowFrom, 0, rows, 0);
			_formulas.Insert(rowFrom, 0, rows, 0);
			_commentsStore.Insert(rowFrom, 0, rows, 0);
			_hyperLinks.Insert(rowFrom, 0, rows, 0);
			_flags.Insert(rowFrom, 0, rows, 0);
			Comments.Insert(rowFrom, 0, rows, 0);
			_names.Insert(rowFrom, 0, rows, 0);
			_sparklineGroups?.Insert(rowFrom, 0, rows, 0);
			Workbook.Names.Insert(rowFrom, 0, rows, 0, n => n.Worksheet == this);

			ShiftSharedFormulasForRowInsert(rowFrom, rows);

			UpdateCellFormulaReferences(rows, 0, rowFrom, 0);

			FixMergedCellsRow(rowFrom, rows, false);
			if (copyStylesFromRow > 0)
			{
				CopyRowStyles(rowFrom, rows, copyStylesFromRow);
			}

			InsertRowInTablesAndValidations(rowFrom, rows);
		}

		// Update cross-sheet references.
		foreach (var sheet in Workbook.Worksheets.Where(sheet => sheet != this))
		{
			sheet.UpdateCrossSheetReferences(Name, rowFrom, rows, 0, 0);
		}
	}

	private void ShiftSharedFormulasForRowInsert(int rowFrom, int rows)
	{
		foreach (var f in _sharedFormulas.Values)
		{
			if (f.StartRow >= rowFrom)
			{
				f.StartRow += rows;
			}

			var a = new ExcelAddressBase(f.Address);
			if (a._fromRow >= rowFrom)
			{
				a._fromRow += rows;
				a._toRow += rows;
			}
			else if (a._toRow >= rowFrom)
			{
				a._toRow += rows;
			}

			f.Address = ExcelCellBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
			f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, 0, Name, Name);
		}
	}

	private void UpdateCellFormulaReferences(int rowIncrement, int colIncrement, int afterRow, int afterColumn)
	{
		var cse = new CellsStoreEnumerator<object>(_formulas);
		while (cse.Next())
		{
			if (cse.Value is string)
			{
				cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), rowIncrement, colIncrement, afterRow, afterColumn, Name, Name);
			}
		}
	}

	private void CopyRowStyles(int rowFrom, int rows, int copyStylesFromRow)
	{
		var cseS = new CellsStoreEnumerator<ExcelCoreValue>(_values, copyStylesFromRow, 0, copyStylesFromRow, ExcelPackage.MaxColumns); //Fixes issue 15068 , 15090
		while (cseS.Next())
		{
			if (cseS.Value._styleId == 0)
			{
				continue;
			}

			for (var r = 0; r < rows; r++)
			{
				SetStyleInner(rowFrom + r, cseS.Column, cseS.Value._styleId);
			}
		}

		var newOutlineLevel = Row(copyStylesFromRow + rows).OutlineLevel;
		for (var r = 0; r < rows; r++)
		{
			Row(rowFrom + r).OutlineLevel = newOutlineLevel;
		}
	}

	private void InsertRowInTablesAndValidations(int rowFrom, int rows)
	{
		foreach (var tbl in Tables)
		{
			tbl.Address = tbl.Address.AddRow(rowFrom, rows);
		}

		foreach (var ptbl in PivotTables)
		{
			ptbl.Address = ptbl.Address.AddRow(rowFrom, rows);
			ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.AddRow(rowFrom, rows).Address;
		}

		//Issue 15573
		foreach (var dv in DataValidations.Cast<ExcelDataValidation>())
		{
			var addr = dv.Address;
			var newAddr = addr.AddRow(rowFrom, rows).Address;
			if (addr.Address != newAddr)
			{
				dv.SetAddress(newAddr);
			}
		}
	}
}
