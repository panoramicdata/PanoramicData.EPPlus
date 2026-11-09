using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Delete the specified column from the worksheet.
	/// </summary>
	/// <param name="column">The column to be deleted</param>
	public void DeleteColumn(int column) => DeleteColumn(column, 1);
	/// <summary>
	/// Delete the specified column from the worksheet.
	/// </summary>
	/// <param name="columnFrom">The start column</param>
	/// <param name="columns">Number of columns to delete</param>
	public void DeleteColumn(int columnFrom, int columns)
	{
		if (columnFrom < 1 || columnFrom + columns > ExcelPackage.MaxColumns)
		{
			throw new ArgumentException("Column out of range. Spans from 1 to " + ExcelPackage.MaxColumns.ToString(CultureInfo.InvariantCulture));
		}

		lock (this)
		{
			TrimColumnBeforeDelete(columnFrom);

			_values.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
			_formulas.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
			_flags.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
			_commentsStore.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
			_hyperLinks.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
			_names.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);

			Comments.Delete(0, columnFrom, 0, columns);
			Workbook.Names.Delete(0, columnFrom, ExcelPackage.MaxRows, columns, n => n.Worksheet == this);

			AdjustFormulasColumn(columnFrom, columns);
			FixMergedCellsColumn(columnFrom, columns, true);

			ShiftColumnDefinitionsForDelete(columnFrom, columns);

			DeleteColumnInTables(columnFrom, columns);

			//Issue 15573
			DeleteColumnInDataValidations(columnFrom, columns);
		}
	}

	private void TrimColumnBeforeDelete(int columnFrom)
	{
		var col = GetValueInner(0, columnFrom) as ExcelColumn;
		if (col == null)
		{
			var r = 0;
			var c = columnFrom;
			if (_values.PrevCell(ref r, ref c))
			{
				col = GetValueInner(0, c) as ExcelColumn;
				if (col._columnMax >= columnFrom)
				{
					col.ColumnMax = columnFrom - 1;
				}
			}
		}
	}

	private void ShiftColumnDefinitionsForDelete(int columnFrom, int columns)
	{
		var csec = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, columnFrom, 0, ExcelPackage.MaxColumns);
		foreach (var val in csec)
		{
			var column = val._value;
			if (column is ExcelColumn c)
			{
				if (c._columnMin >= columnFrom)
				{
					c._columnMin -= columns;
					c._columnMax -= columns;
				}
			}
		}
	}

}
