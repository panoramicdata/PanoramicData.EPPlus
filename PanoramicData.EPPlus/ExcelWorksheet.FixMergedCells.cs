using OfficeOpenXml.Table;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private static void InsertTableColumns(int columnFrom, int columns, ExcelTable tbl)
	{
		var node = tbl.Columns[0].TopNode.ParentNode;
		var ix = columnFrom - tbl.Address.Start.Column - 1;
		var insPos = node.ChildNodes[ix];
		ix += 2;
		for (var i = 0; i < columns; i++)
		{
			var name =
				tbl.Columns.GetUniqueName(string.Format("Column{0}",
					(ix++).ToString(CultureInfo.InvariantCulture)));
			var tableColumn =
				(XmlElement)tbl.TableXml.CreateNode(XmlNodeType.Element, "tableColumn", ExcelPackage.schemaMain);
			tableColumn.SetAttribute("id", (tbl.Columns.Count + i + 1).ToString(CultureInfo.InvariantCulture));
			tableColumn.SetAttribute("name", name);
			insPos = node.InsertAfter(tableColumn, insPos);
		} //Create tbl Column

		tbl._cols = new ExcelTableColumnCollection(tbl);
	}

	/// <summary>
	/// Adds a value to the row of merged cells to fix for inserts or deletes
	/// </summary>
	/// <param name="row"></param>
	/// <param name="rows"></param> 
	/// <param name="delete"></param>
	private void FixMergedCellsRow(int row, int rows, bool delete)
	{
		if (delete)
		{
			_mergedCells._cells.Delete(row, 0, rows, 0);
		}
		else
		{
			_mergedCells._cells.Insert(row, 0, rows, 0);
		}

		FixMergedCellAddresses(delete, addr => delete ? addr.DeleteRow(row, rows) : addr.AddRow(row, rows));
	}
	/// <summary>
	/// Adds a value to the row of merged cells to fix for inserts or deletes
	/// </summary>
	/// <param name="column"></param>
	/// <param name="columns"></param>
	/// <param name="delete"></param>
	private void FixMergedCellsColumn(int column, int columns, bool delete)
	{
		if (delete)
		{
			_mergedCells._cells.Delete(0, column, 0, columns);
		}
		else
		{
			_mergedCells._cells.Insert(0, column, 0, columns);
		}

		FixMergedCellAddresses(delete, addr => delete ? addr.DeleteColumn(column, columns) : addr.AddColumn(column, columns));
	}

	private void FixMergedCellAddresses(bool delete, Func<ExcelAddressBase, ExcelAddressBase> shift)
	{
		List<int> removeIndex = [];
		for (var i = 0; i < _mergedCells.Count; i++)
		{
			if (!string.IsNullOrEmpty(_mergedCells[i]))
			{
				ExcelAddressBase addr = new(_mergedCells[i]);
				var newAddr = shift(addr);
				if (delete)
				{
					if (newAddr == null)
					{
						removeIndex.Add(i);
						continue;
					}
				}
				else if (newAddr.Address != addr.Address)
				{
					_mergedCells.SetIndex(newAddr, i);
				}

				if (newAddr.Address != addr.Address)
				{
					_mergedCells.List[i] = newAddr._address;
				}
			}
		}

		for (var i = removeIndex.Count - 1; i >= 0; i--)
		{
			_mergedCells.List.RemoveAt(removeIndex[i]);
		}
	}
}
