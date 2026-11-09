using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	internal void SetSharedFormulaID(int id)
	{
		for (var col = _fromCol; col <= _toCol; col++)
		{
			for (var row = _fromRow; row <= _toRow; row++)
			{
				_worksheet._formulas.SetValue(row, col, id);
			}
		}
	}
	private void CheckAndSplitSharedFormula(ExcelAddressBase address)
	{
		for (var col = address._fromCol; col <= address._toCol; col++)
		{
			for (var row = address._fromRow; row <= address._toRow; row++)
			{
				var f = _worksheet._formulas.GetValue(row, col);
				if (f is int v && v >= 0)
				{
					SplitFormulas(address);
					return;
				}
			}
		}
	}

	private void SplitFormulas(ExcelAddressBase address)
	{
		List<int> formulas = [];
		for (var col = address._fromCol; col <= address._toCol; col++)
		{
			for (var row = address._fromRow; row <= address._toRow; row++)
			{
				var f = _worksheet._formulas.GetValue(row, col);
				if (f is int id)
				{
					if (id >= 0 && !formulas.Contains(id))
					{
						if (_worksheet._sharedFormulas[id].IsArray &&
								Collide(_worksheet.Cells[_worksheet._sharedFormulas[id].Address]) == eAddressCollition.Partly) // If the formula is an array formula and its on the inside the overwriting range throw an exception
						{
							throw new InvalidOperationException("Can not overwrite a part of an array-formula");
						}

						formulas.Add(id);
					}
				}
			}
		}

		foreach (var ix in formulas)
		{
			SplitFormula(address, ix);
		}

	}

	private void SplitFormula(ExcelAddressBase address, int ix)
	{
		var f = _worksheet._sharedFormulas[ix];
		var fRange = _worksheet.Cells[f.Address];
		var collide = address.Collide(fRange);

		//The formula is inside the currenct range, remove it
		if (collide is eAddressCollition.Equal or eAddressCollition.Inside)
		{
			_worksheet._sharedFormulas.Remove(ix);
			return;
		}

		var firstCellCollide = address.Collide(new ExcelAddressBase(fRange._fromRow, fRange._fromCol, fRange._fromRow, fRange._fromCol));
		if (collide == eAddressCollition.Partly && (firstCellCollide == eAddressCollition.Inside || firstCellCollide == eAddressCollition.Equal)) //Do we need to split? Only if the functions first row is inside the new range.
		{
			//The formula partly collides with the current range
			var fIsSet = false;
			var formulaR1C1 = fRange.FormulaR1C1;
			//Top Range
			if (fRange._fromRow < _fromRow)
			{
				f.Address = GetAddress(fRange._fromRow, fRange._fromCol, _fromRow - 1, fRange._toCol);
				fIsSet = true;
			}
			//Left Range
			if (fRange._fromCol < address._fromCol)
			{
				if (fIsSet)
				{
					f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default)
					{
						Index = _worksheet.GetMaxShareFunctionIndex(false),
						StartCol = fRange._fromCol,
						IsArray = false
					};
					_worksheet._sharedFormulas.Add(f.Index, f);
				}
				else
				{
					fIsSet = true;
				}

				if (fRange._fromRow < address._fromRow)
				{
					f.StartRow = address._fromRow;
				}
				else
				{
					f.StartRow = fRange._fromRow;
				}

				if (fRange._toRow < address._toRow)
				{
					f.Address = GetAddress(f.StartRow, f.StartCol,
							fRange._toRow, address._fromCol - 1);
				}
				else
				{
					f.Address = GetAddress(f.StartRow, f.StartCol,
						 address._toRow, address._fromCol - 1);
				}

				f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);
				_worksheet.Cells[f.Address].SetSharedFormulaID(f.Index);
			}
			//Right Range
			if (fRange._toCol > address._toCol)
			{
				if (fIsSet)
				{
					f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default)
					{
						Index = _worksheet.GetMaxShareFunctionIndex(false),
						IsArray = false
					};
					_worksheet._sharedFormulas.Add(f.Index, f);
				}
				else
				{
					fIsSet = true;
				}

				f.StartCol = address._toCol + 1;
				if (address._fromRow < fRange._fromRow)
				{
					f.StartRow = fRange._fromRow;
				}
				else
				{
					f.StartRow = address._fromRow;
				}

				if (fRange._toRow < address._toRow)
				{
					f.Address = GetAddress(f.StartRow, f.StartCol,
							fRange._toRow, fRange._toCol);
				}
				else
				{
					f.Address = GetAddress(f.StartRow, f.StartCol,
							address._toRow, fRange._toCol);
				}

				f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);
				_worksheet.Cells[f.Address].SetSharedFormulaID(f.Index);
			}
			//Bottom Range
			if (fRange._toRow > address._toRow)
			{
				if (fIsSet)
				{
					f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default)
					{
						Index = _worksheet.GetMaxShareFunctionIndex(false),
						IsArray = false
					};
					_worksheet._sharedFormulas.Add(f.Index, f);
				}

				f.StartCol = fRange._fromCol;
				f.StartRow = _toRow + 1;

				f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);

				f.Address = GetAddress(f.StartRow, f.StartCol,
						fRange._toRow, fRange._toCol);
				_worksheet.Cells[f.Address].SetSharedFormulaID(f.Index);

			}
		}
	}
}
