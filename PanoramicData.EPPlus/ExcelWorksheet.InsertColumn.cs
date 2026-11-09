using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Table.PivotTable;

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

			var cse = new CellsStoreEnumerator<object>(_formulas);
			while (cse.Next())
			{
				if (cse.Value is string)
				{
					cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), 0, columns, 0, columnFrom, Name, Name);
				}
			}

			FixMergedCellsColumn(columnFrom, columns, false);

			var csec = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 1, 0, ExcelPackage.MaxColumns);
			var lst = new List<ExcelColumn>();
			foreach (var val in csec)
			{
				var col = val._value;
				if (col is ExcelColumn column)
				{
					lst.Add(column);
				}
			}

			for (var i = lst.Count - 1; i >= 0; i--)
			{
				var c = lst[i];
				if (c._columnMin >= columnFrom)
				{
					if (c._columnMin + columns <= ExcelPackage.MaxColumns)
					{
						c._columnMin += columns;
					}
					else
					{
						c._columnMin = ExcelPackage.MaxColumns;
					}

					if (c._columnMax + columns <= ExcelPackage.MaxColumns)
					{
						c._columnMax += columns;
					}
					else
					{
						c._columnMax = ExcelPackage.MaxColumns;
					}
				}
				else if (c._columnMax >= columnFrom)
				{
					var cc = c._columnMax - columnFrom;
					c._columnMax = columnFrom - 1;
					CopyColumn(c, columnFrom + columns, columnFrom + columns + cc);
				}
			}


			//Copy style from another column?
			if (copyStylesFromColumn > 0)
			{
				if (copyStylesFromColumn >= columnFrom)
				{
					copyStylesFromColumn += columns;
				}

				//Get styles to a cached list, 
				var l = new List<int[]>();
				var sce = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, copyStylesFromColumn, ExcelPackage.MaxRows, copyStylesFromColumn);
				lock (sce)
				{
					while (sce.Next())
					{
						if (sce.Value._styleId == 0)
						{
							continue;
						}

						l.Add([sce.Row, sce.Value._styleId]);
					}
				}

				//Set the style id's from the list.
				foreach (var sc in l)
				{
					for (var c = 0; c < columns; c++)
					{
						if (sc[0] == 0)
						{
							var col = Column(columnFrom + c);   //Create the column
							col.StyleID = sc[1];
						}
						else
						{
							SetStyleInner(sc[0], columnFrom + c, sc[1]);
						}
					}
				}

				var newOutlineLevel = Column(copyStylesFromColumn).OutlineLevel;
				for (var c = 0; c < columns; c++)
				{
					Column(columnFrom + c).OutlineLevel = newOutlineLevel;
				}
			}
			//Adjust tables
			foreach (var tbl in Tables)
			{
				if (columnFrom > tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
				{
					InsertTableColumns(columnFrom, columns, tbl);
				}

				tbl.Address = tbl.Address.AddColumn(columnFrom, columns);
			}

			foreach (var ptbl in PivotTables)
			{
				if (columnFrom <= ptbl.Address.End.Column)
				{
					ptbl.Address = ptbl.Address.AddColumn(columnFrom, columns);
				}

				if (columnFrom <= ptbl.CacheDefinition.SourceRange.End.Column)
				{
					if (ptbl.CacheDefinition.CacheSource == eSourceType.Worksheet)
					{
						ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.AddColumn(columnFrom, columns).Address;
					}
				}
			}

			//Issue 15573
			foreach (var dv in DataValidations.Cast<ExcelDataValidation>())
			{
				var addr = dv.Address;
				var newAddr = addr.AddColumn(columnFrom, columns).Address;
				if (addr.Address != newAddr)
				{
					dv.SetAddress(newAddr);
				}
			}

			// Update cross-sheet references.
			foreach (var sheet in Workbook.Worksheets.Where(sheet => sheet != this))
			{
				sheet.UpdateCrossSheetReferences(Name, 0, 0, columnFrom, columns);
			}
		}
	}
}
