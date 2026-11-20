using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	internal void AdjustFormulasRow(int rowFrom, int rows)
	{
		var delSF = new List<int>();
		foreach (var sf in _sharedFormulas.Values)
		{
			var a = new ExcelAddress(sf.Address).DeleteRow(rowFrom, rows);
			if (a == null)
			{
				delSF.Add(sf.Index);
			}
			else
			{
				sf.Address = a.Address;
				if (sf.StartRow > rowFrom)
				{
					var r = Math.Min(sf.StartRow - rowFrom, rows);
					sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, -r, 0, rowFrom, 0, Name, Name);
					sf.StartRow -= r;
				}
			}
		}

		RemoveSharedFormulasAndUpdateReferences(delSF, -rows, 0, rowFrom, 0);
	}

	private void RemoveSharedFormulasAndUpdateReferences(List<int> delSF, int rowIncrement, int colIncrement, int afterRow, int afterColumn)
	{
		foreach (var ix in delSF)
		{
			_sharedFormulas.Remove(ix);
		}

		var cse = new CellsStoreEnumerator<object>(_formulas, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
		while (cse.Next())
		{
			if (cse.Value is string)
			{
				cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), rowIncrement, colIncrement, afterRow, afterColumn, Name, Name);
			}
		}
	}
	internal void AdjustFormulasColumn(int columnFrom, int columns)
	{
		var delSF = new List<int>();
		foreach (var sf in _sharedFormulas.Values)
		{
			var a = new ExcelAddress(sf.Address).DeleteColumn(columnFrom, columns);
			if (a == null)
			{
				delSF.Add(sf.Index);
			}
			else
			{
				sf.Address = a.Address;
				if (sf.StartCol > columnFrom)
				{
					var c = Math.Min(sf.StartCol - columnFrom, columns);
					sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, 0, -c, 0, 1, Name, Name);
					sf.StartCol -= c;
				}
			}
		}

		RemoveSharedFormulasAndUpdateReferences(delSF, 0, -columns, 0, columnFrom);
	}


	[SuppressMessage("Blocker Bug", "S2551:Shared resources should not be used for locking", Justification = "Behaviour-preserving: the worksheet instance is the lock used by the other structural edits, so changing it would break their mutual exclusion.")]
	private void UpdateCrossSheetReferences(string sheetWhoseReferencesShouldBeUpdated, int rowFrom, int rows, int columnFrom, int columns)
	{
		lock (this)
		{
			foreach (var f in _sharedFormulas.Values)
			{
				f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, columns, rowFrom, columnFrom, Name, sheetWhoseReferencesShouldBeUpdated);
			}

			var cse = new CellsStoreEnumerator<object>(_formulas);
			while (cse.Next())
			{
				if (cse.Value is string)
				{
					cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), rows, columns, rowFrom, columnFrom, Name, sheetWhoseReferencesShouldBeUpdated);
				}
			}
		}
	}

	private void UpdateCrossSheetReferenceNames(string oldName, string newName)
	{
		if (string.IsNullOrEmpty(oldName))
		{
			throw new ArgumentNullException(nameof(oldName));
		}

		if (string.IsNullOrEmpty(newName))
		{
			throw new ArgumentNullException(nameof(newName));
		}

		lock (this)
		{
			foreach (var f in _sharedFormulas.Values)
			{
				f.Formula = ExcelCellBase.UpdateFormulaSheetReferences(f.Formula, oldName, newName);
			}

			var cse = new CellsStoreEnumerator<object>(_formulas);
			while (cse.Next())
			{
				if (cse.Value is string)
				{
					cse.Value = ExcelCellBase.UpdateFormulaSheetReferences(cse.Value.ToString(), oldName, newName);
				}
			}
		}
	}
}
