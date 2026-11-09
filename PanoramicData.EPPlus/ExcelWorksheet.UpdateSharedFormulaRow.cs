using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void UpdateSharedFormulaRow(ref Formulas formula, int startRow, int rows, ref List<Formulas> newFormulas)
	{
		var newFormulasCount = newFormulas.Count;
		ExcelCellBase.GetRowColFromAddress(formula.Address, out var fromRow, out var fromCol, out var toRow, out int toCol);
		var formualR1C1 = TranslateSharedFormulaStart(formula, startRow, rows, fromRow);
		var prevFormualR1C1 = formualR1C1;
		for (var row = fromRow; row <= toRow; row++)
		{
			for (var col = fromCol; col <= toCol; col++)
			{
				var newFormula = GetShiftedCellFormula(formualR1C1, row, col, startRow, rows, out var currentFormulaR1C1);

				if (currentFormulaR1C1 != prevFormualR1C1)
				{
					if (row == fromRow && col == fromCol)
					{
						formula.Formula = newFormula;
					}
					else
					{
						SplitSharedFormula(formula, newFormulas, newFormulasCount, row, col, newFormula);

						prevFormualR1C1 = currentFormulaR1C1;
					}
				}
			}
		}

		AdjustSharedFormulaStartRow(formula, startRow, rows);

		if (newFormulas.Count > newFormulasCount)
		{
			newFormulas[^1].Address = ExcelCellBase.GetAddress(newFormulas[^1].StartRow, newFormulas[^1].StartCol, toRow, toCol);
		}
	}

	private static string TranslateSharedFormulaStart(Formulas formula, int startRow, int rows, int fromRow)
	{
		string formualR1C1;
		if (rows > 0 || fromRow <= startRow)
		{
			formualR1C1 = ExcelCellBase.TranslateToR1C1(formula.Formula, formula.StartRow, formula.StartCol);
			formula.Formula = ExcelCellBase.TranslateFromR1C1(formualR1C1, fromRow, formula.StartCol);
		}
		else
		{
			formualR1C1 = ExcelCellBase.TranslateToR1C1(formula.Formula, formula.StartRow - rows, formula.StartCol);
			formula.Formula = ExcelCellBase.TranslateFromR1C1(formualR1C1, formula.StartRow, formula.StartCol);
		}

		return formualR1C1;
	}

	private string GetShiftedCellFormula(string formualR1C1, int row, int col, int startRow, int rows, out string currentFormulaR1C1)
	{
		string newFormula;
		if (rows > 0 || row < startRow)
		{
			newFormula = ExcelCellBase.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(formualR1C1, row, col), rows, 0, startRow, 0, Name, Name);
			currentFormulaR1C1 = ExcelCellBase.TranslateToR1C1(newFormula, row, col);
		}
		else
		{
			newFormula = ExcelCellBase.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(formualR1C1, row - rows, col), rows, 0, startRow, 0, Name, Name);
			currentFormulaR1C1 = ExcelCellBase.TranslateToR1C1(newFormula, row, col);
		}

		return newFormula;
	}

	private static void SplitSharedFormula(Formulas formula, List<Formulas> newFormulas, int newFormulasCount, int row, int col, string newFormula)
	{
		if (newFormulas.Count == newFormulasCount)
		{
			formula.Address = ExcelCellBase.GetAddress(formula.StartRow, formula.StartCol, row - 1, col);
		}
		else
		{
			newFormulas[^1].Address = ExcelCellBase.GetAddress(newFormulas[^1].StartRow, newFormulas[^1].StartCol, row - 1, col);
		}

		var refFormula = new Formulas(SourceCodeTokenizer.Default)
		{
			Formula = newFormula,
			StartRow = row,
			StartCol = col
		};
		newFormulas.Add(refFormula);
	}

	private static void AdjustSharedFormulaStartRow(Formulas formula, int startRow, int rows)
	{
		if (rows < 0 && formula.StartRow > startRow)
		{
			if (formula.StartRow + rows < startRow)
			{
				formula.StartRow = startRow;
			}
			else
			{
				formula.StartRow += rows;
			}
		}
	}
}
