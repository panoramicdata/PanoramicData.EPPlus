using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void InsertColumnInTables(int columnFrom, int columns)
	{
		foreach (var tbl in Tables)
		{
			if (columnFrom > tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
			{
				InsertTableColumns(columnFrom, columns, tbl);
			}

			tbl.Address = tbl.Address.AddColumn(columnFrom, columns);
		}
	}

	private void InsertColumnInPivotTables(int columnFrom, int columns)
	{
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
	}

	private void InsertColumnInDataValidations(int columnFrom, int columns)
	{
		foreach (var dv in DataValidations.Cast<ExcelDataValidation>())
		{
			var addr = dv.Address;
			var newAddr = addr.AddColumn(columnFrom, columns).Address;
			if (addr.Address != newAddr)
			{
				dv.SetAddress(newAddr);
			}
		}
	}
}