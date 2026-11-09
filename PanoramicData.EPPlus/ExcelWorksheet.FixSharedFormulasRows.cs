using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void FixSharedFormulasRows(int position, int rows)
	{
		List<Formulas> added = [];
		List<Formulas> deleted = [];

		foreach (var id in _sharedFormulas.Keys)
		{
			var f = _sharedFormulas[id];

			ExcelCellBase.GetRowColFromAddress(f.Address, out var fromRow, out _, out int toRow, out _);
			if (position >= fromRow && position + (Math.Abs(rows)) <= toRow) //Insert/delete is whithin the share formula address
			{
				ResizeSharedFormulaWithin(f, position, rows, added);
			}
			else if (position <= toRow)
			{
				ShiftSharedFormulaAfter(f, position, rows, deleted);
			}
		}

		AddFormulas(added);

		//Remove formulas
		foreach (var f in deleted)
		{
			_sharedFormulas.Remove(f.Index);
		}

		//Fix Formulas
		added = [];
		foreach (var id in _sharedFormulas.Keys)
		{
			var f = _sharedFormulas[id];
			UpdateSharedFormulaRow(ref f, position, rows, ref added);
		}

		AddFormulas(added);
	}

	private static void ResizeSharedFormulaWithin(Formulas f, int position, int rows, List<Formulas> added)
	{
		ExcelCellBase.GetRowColFromAddress(f.Address, out var fromRow, out var fromCol, out int toRow, out var toCol);
		if (rows > 0) //Insert
		{
			f.Address = ExcelCellBase.GetAddress(fromRow, fromCol) + ":" + ExcelCellBase.GetAddress(position - 1, toCol);
			if (toRow != fromRow)
			{
				Formulas newF = new(SourceCodeTokenizer.Default)
				{
					StartCol = f.StartCol,
					StartRow = position + rows,
					Address = ExcelCellBase.GetAddress(position + rows, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol),
					Formula = ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), position, f.StartCol)
				};
				added.Add(newF);
			}
		}
		else
		{
			if (fromRow - rows < toRow)
			{
				f.Address = ExcelCellBase.GetAddress(fromRow, fromCol, toRow + rows, toCol);
			}
			else
			{
				f.Address = ExcelCellBase.GetAddress(fromRow, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol);
			}
		}
	}

	private void ShiftSharedFormulaAfter(Formulas f, int position, int rows, List<Formulas> deleted)
	{
		ExcelCellBase.GetRowColFromAddress(f.Address, out var fromRow, out var fromCol, out int toRow, out var toCol);
		if (rows > 0) //Insert before shift down
		{
			f.StartRow += rows;
			f.Address = ExcelCellBase.GetAddress(fromRow + rows, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol);
		}
		else
		{
			if (position <= fromRow && position + Math.Abs(rows) > toRow)  //Delete the formula 
			{
				deleted.Add(f);
			}
			else
			{
				toRow = toRow + rows < position - 1 ? position - 1 : toRow + rows;
				if (position <= fromRow)
				{
					fromRow = fromRow + rows < position ? position : fromRow + rows;
				}

				f.Address = ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol);
				Cells[f.Address].SetSharedFormulaID(f.Index);
			}
		}
	}

	private void AddFormulas(List<Formulas> added)
	{
		//Add new formulas
		foreach (var f in added)
		{
			f.Index = GetMaxShareFunctionIndex(false);
			_sharedFormulas.Add(f.Index, f);
			Cells[f.Address].SetSharedFormulaID(f.Index);
		}
	}
}
