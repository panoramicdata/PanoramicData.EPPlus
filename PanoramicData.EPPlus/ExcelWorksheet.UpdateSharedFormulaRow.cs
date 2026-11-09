using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void UpdateSharedFormulaRow(ref Formulas formula, int startRow, int rows, ref List<Formulas> newFormulas)
	{
		var newFormulasCount = newFormulas.Count;
		ExcelCellBase.GetRowColFromAddress(formula.Address, out var fromRow, out var fromCol, out var toRow, out int toCol);
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
		var prevFormualR1C1 = formualR1C1;
		for (var row = fromRow; row <= toRow; row++)
		{
			for (var col = fromCol; col <= toCol; col++)
			{
				string newFormula;
				string currentFormulaR1C1;
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

				if (currentFormulaR1C1 != prevFormualR1C1)
				{
					if (row == fromRow && col == fromCol)
					{
						formula.Formula = newFormula;
					}
					else
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

						prevFormualR1C1 = currentFormulaR1C1;
					}
				}
			}
		}

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

		if (newFormulas.Count > newFormulasCount)
		{
			newFormulas[^1].Address = ExcelCellBase.GetAddress(newFormulas[^1].StartRow, newFormulas[^1].StartCol, toRow, toCol);
		}
	}
}
