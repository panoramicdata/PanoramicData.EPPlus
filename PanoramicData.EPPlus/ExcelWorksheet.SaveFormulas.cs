using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Check all Shared formulas that the first cell has not been deleted.
	/// If so create a standard formula of all cells in the formula .
	/// </summary>
	private void FixSharedFormulas()
	{
		var remove = new List<int>();
		foreach (var f in _sharedFormulas.Values)
		{
			var addr = new ExcelAddressBase(f.Address);
			var shIx = _formulas.GetValue(addr._fromRow, addr._fromCol);
			if (shIx is not int || (shIx is int v && v != f.Index))
			{
				ReplaceSharedFormulaWithCellFormulas(f, addr);

				remove.Add(f.Index);
			}
		}

		remove.ForEach(i => _sharedFormulas.Remove(i));
	}

	private void ReplaceSharedFormulaWithCellFormulas(Formulas f, ExcelAddressBase addr)
	{
		for (var row = addr._fromRow; row <= addr._toRow; row++)
		{
			for (var col = addr._fromCol; col <= addr._toCol; col++)
			{
				if (!(addr._fromRow == row && addr._fromCol == col))
				{
					var fIx = _formulas.GetValue(row, col);
					if (fIx is int v1 && v1 == f.Index)
					{
						_formulas.SetValue(row, col, f.GetFormula(row, col, Name));
					}
				}
			}
		}
	}

	private object GetFormulaValue(object v)
	{
		if (v != null && v.ToString() != "")
		{
			return "<v>" + ConvertUtil.ExcelEscapeString(GetValueForXml(v)) + "</v>"; //Fixes issue 15071
		}
		else
		{
			return "";
		}
	}
}
