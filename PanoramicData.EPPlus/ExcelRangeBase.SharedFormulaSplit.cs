using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private void SplitPartlyCollidingFormula(ExcelAddressBase address, ExcelWorksheet.Formulas formula, ExcelRange fRange)
	{
		var f = formula;
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
			f = SplitFormulaLeftRange(address, f, fRange, formulaR1C1, ref fIsSet);
		}
		//Right Range
		if (fRange._toCol > address._toCol)
		{
			f = SplitFormulaRightRange(address, f, fRange, formulaR1C1, ref fIsSet);
		}
		//Bottom Range
		if (fRange._toRow > address._toRow)
		{
			SplitFormulaBottomRange(f, fRange, formulaR1C1, fIsSet);
		}
	}

	private ExcelWorksheet.Formulas SplitFormulaLeftRange(ExcelAddressBase address, ExcelWorksheet.Formulas current, ExcelRange fRange, string formulaR1C1, ref bool fIsSet)
	{
		var f = current;
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
		return f;
	}

	private ExcelWorksheet.Formulas SplitFormulaRightRange(ExcelAddressBase address, ExcelWorksheet.Formulas current, ExcelRange fRange, string formulaR1C1, ref bool fIsSet)
	{
		var f = current;
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
		return f;
	}

	private void SplitFormulaBottomRange(ExcelWorksheet.Formulas current, ExcelRange fRange, string formulaR1C1, bool fIsSet)
	{
		var f = current;
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
