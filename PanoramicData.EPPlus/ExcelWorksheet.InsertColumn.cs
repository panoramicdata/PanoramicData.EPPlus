namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Inserts a new column into the spreadsheet.  Existing columns below the position are 
	/// shifted down.  All formula are updated to take account of the new column.
	/// </summary>
	/// <param name="columnFrom">The position of the new column</param>
	/// <param name="columns">Number of columns to insert</param>        
	public void InsertColumn(int columnFrom, int columns) => InsertColumn(columnFrom, columns, 0);
	///<summary>
	/// Inserts a new column into the spreadsheet.  Existing column to the left are 
	/// shifted.  All formula are updated to take account of the new column.
	/// </summary>
	/// <param name="columnFrom">The position of the new column</param>
	/// <param name="columns">Number of columns to insert.</param>
	/// <param name="copyStylesFromColumn">Copy Styles from this column. Applied to all inserted columns</param>
	public void InsertColumn(int columnFrom, int columns, int copyStylesFromColumn)
	{
		CheckSheetType();
		var d = Dimension;

		if (columnFrom < 1)
		{
			throw new ArgumentOutOfRangeException("columnFrom can't be lesser that 1");
		}
		//Check that cells aren't shifted outside the boundries
		if (d != null && d.End.Column > columnFrom && d.End.Column + columns > ExcelPackage.MaxColumns)
		{
			throw new ArgumentOutOfRangeException("Can't insert. Columns will be shifted outside the boundries of the worksheet.");
		}

		lock (this)
		{
			_values.Insert(0, columnFrom, 0, columns);
			_formulas.Insert(0, columnFrom, 0, columns);
			_commentsStore.Insert(0, columnFrom, 0, columns);
			_hyperLinks.Insert(0, columnFrom, 0, columns);
			_flags.Insert(0, columnFrom, 0, columns);
			_names.Insert(0, columnFrom, 0, columns);
			_sparklineGroups?.Insert(0, columnFrom, 0, columns);
			Comments.Insert(0, columnFrom, 0, columns);
			Workbook.Names.Insert(0, columnFrom, 0, columns, n => n.Worksheet == this);

			ShiftSharedFormulasForColumnInsert(columnFrom, columns);

			UpdateCellFormulaReferences(0, columns, 0, columnFrom);

			FixMergedCellsColumn(columnFrom, columns, false);

			ShiftColumnDefinitions(columnFrom, columns);

			//Copy style from another column?
			if (copyStylesFromColumn > 0)
			{
				CopyColumnStyles(columnFrom, columns, copyStylesFromColumn);
			}
			//Adjust tables
			InsertColumnInTables(columnFrom, columns);

			InsertColumnInPivotTables(columnFrom, columns);

			//Issue 15573
			InsertColumnInDataValidations(columnFrom, columns);

			// Update cross-sheet references.
			foreach (var sheet in Workbook.Worksheets.Where(sheet => sheet != this))
			{
				sheet.UpdateCrossSheetReferences(Name, 0, 0, columnFrom, columns);
			}
		}
	}

	private void ShiftSharedFormulasForColumnInsert(int columnFrom, int columns)
	{
		foreach (var f in _sharedFormulas.Values)
		{
			if (f.StartCol >= columnFrom)
			{
				f.StartCol += columns;
			}

			var a = new ExcelAddressBase(f.Address);
			if (a._fromCol >= columnFrom)
			{
				a._fromCol += columns;
				a._toCol += columns;
			}
			else if (a._toCol >= columnFrom)
			{
				a._toCol += columns;
			}

			f.Address = ExcelCellBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
			f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, 0, columns, 0, columnFrom, Name, Name);
		}
	}
}
