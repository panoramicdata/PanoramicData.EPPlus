using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Table;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
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
