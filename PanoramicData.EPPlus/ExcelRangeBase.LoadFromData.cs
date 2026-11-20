using OfficeOpenXml.Table;
using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Load the data from the datareader starting from the top left cell of the range
	/// </summary>
	/// <param name="Reader">The datareader to loadfrom</param>
	/// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
	/// <param name="TableName">The name of the table</param>
	/// <param name="TableStyle">The table style to apply to the data</param>
	/// <returns>The filled range</returns>
	[SuppressMessage("Major Code Smell", "S2360:Optional parameters should not be used", Justification = "Public API: replacing the optional parameter with overloads would break binary compatibility.")]
	public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders, string TableName, TableStyles TableStyle = TableStyles.None)
	{
		var r = LoadFromDataReader(Reader, PrintHeaders);

		var rows = r.Rows - 1;
		if (rows >= 0 && r.Columns > 0)
		{
			var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + (rows <= 0 ? 1 : rows), _fromCol + r.Columns - 1), TableName);
			tbl.ShowHeader = PrintHeaders;
			tbl.TableStyle = TableStyle;
		}

		return r;
	}

	/// <summary>
	/// Load the data from the datareader starting from the top left cell of the range
	/// </summary>
	/// <param name="Reader">The datareader to load from</param>
	/// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
	/// <returns>The filled range</returns>
	public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders)
	{
		if (Reader == null)
		{
			throw new ArgumentNullException(nameof(Reader), "Reader can't be null");
		}

		var fieldCount = Reader.FieldCount;

		int col = _fromCol, row = _fromRow;
		if (PrintHeaders)
		{
			for (var i = 0; i < fieldCount; i++)
			{
				// If no caption is set, the ColumnName property is called implicitly.
				_worksheet.SetValueInner(row, col++, Reader.GetName(i));
			}

			row++;
			col = _fromCol;
		}

		while (Reader.Read())
		{
			for (var i = 0; i < fieldCount; i++)
			{
				_worksheet.SetValueInner(row, col++, Reader.GetValue(i));
			}

			row++;
			col = _fromCol;
		}

		return _worksheet.Cells[_fromRow, _fromCol, row - 1, _fromCol + fieldCount - 1];
	}

	/// <summary>
	/// Load the data from the datatable starting from the top left cell of the range
	/// </summary>
	/// <param name="Table">The datatable to load</param>
	/// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
	/// <param name="TableStyle">The table style to apply to the data</param>
	/// <returns>The filled range</returns>
	public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders, TableStyles TableStyle)
	{
		var r = LoadFromDataTable(Table, PrintHeaders);

		var rows = (Table.Rows.Count == 0 ? 1 : Table.Rows.Count) + (PrintHeaders ? 1 : 0);
		if (rows >= 0 && Table.Columns.Count > 0)
		{
			var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + rows - 1, _fromCol + Table.Columns.Count - 1), Table.TableName);
			tbl.ShowHeader = PrintHeaders;
			tbl.TableStyle = TableStyle;
		}

		return r;
	}
	/// <summary>
	/// Load the data from the datatable starting from the top left cell of the range
	/// </summary>
	/// <param name="Table">The datatable to load</param>
	/// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
	/// <returns>The filled range</returns>
	public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders)
	{
		ArgumentNullException.ThrowIfNull(Table);

		if (Table.Rows.Count == 0 && PrintHeaders == false)
		{
			return null;
		}

		var rowArray = new List<object[]>();
		if (PrintHeaders)
		{
			rowArray.Add(Table.Columns.Cast<DataColumn>().Select((dc) => { return (object)dc.Caption; }).ToArray());
		}

		foreach (DataRow dr in Table.Rows)
		{
			rowArray.Add(dr.ItemArray);
		}

		_worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + Table.Columns.Count - 1,
			(List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value) =>
			{
				rowIx -= _fromRow;
				columnIx -= _fromCol;

				var val = ((List<object[]>)value)[rowIx][columnIx];
				if (val != null && val != DBNull.Value && !string.IsNullOrEmpty(val.ToString()))
				{
					list[index] = new ExcelCoreValue { _value = val, _styleId = list[index]._styleId };
				}
			}, rowArray);

		return _worksheet.Cells[_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + Table.Columns.Count - 1];
	}
}
