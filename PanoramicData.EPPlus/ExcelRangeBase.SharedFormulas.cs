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
			SplitPartlyCollidingFormula(address, f, fRange);
		}
	}
}
