using OfficeOpenXml.DataValidation;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
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
