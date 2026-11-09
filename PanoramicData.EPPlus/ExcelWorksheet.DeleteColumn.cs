using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Table;
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

	private void DeleteColumnInTables(int columnFrom, int columns)
	{
		foreach (var tbl in Tables)
		{
			if (columnFrom >= tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
			{
				RemoveTableColumnNodes(tbl, columnFrom, columns);
			}

			tbl.Address = tbl.Address.DeleteColumn(columnFrom, columns);

			foreach (var ptbl in PivotTables)
			{
				if (ptbl.Address.Start.Column > columnFrom + columns)
				{
					ptbl.Address = ptbl.Address.DeleteColumn(columnFrom, columns);
				}

				if (ptbl.CacheDefinition.SourceRange.Start.Column > columnFrom + columns)
				{
					ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.DeleteColumn(columnFrom, columns).Address;
				}
			}
		}
	}

	private static void RemoveTableColumnNodes(ExcelTable tbl, int columnFrom, int columns)
	{
		var node = tbl.Columns[0].TopNode.ParentNode;
		var ix = columnFrom - tbl.Address.Start.Column;
		for (var i = 0; i < columns; i++)
		{
			if (node.ChildNodes.Count > ix)
			{
				node.RemoveChild(node.ChildNodes[ix]);
			}
		}

		tbl._cols = new ExcelTableColumnCollection(tbl);
	}

	private void DeleteColumnInDataValidations(int columnFrom, int columns)
	{
		foreach (var dv in DataValidations.Cast<ExcelDataValidation>())
		{
			var addr = dv.Address;
			if (addr.Start.Column > columnFrom + columns)
			{
				var newAddr = addr.DeleteColumn(columnFrom, columns).Address;
				if (addr.Address != newAddr)
				{
					dv.SetAddress(newAddr);
				}
			}
		}
	}
}
