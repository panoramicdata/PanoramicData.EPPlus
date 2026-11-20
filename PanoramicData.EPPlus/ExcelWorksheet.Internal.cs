namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Get the next ID from a shared formula or an Array formula
	/// Sharedforumlas will have an id from 0-x. Array formula ids start from 0x4000001-. 
	/// </summary>
	/// <param name="isArray">If the formula is an array formula</param>
	/// <returns></returns>
	internal int GetMaxShareFunctionIndex(bool isArray)
	{
		var i = _sharedFormulas.Count + 1;
		if (isArray)
		{
			i |= 0x40000000;
		}

		while (_sharedFormulas.ContainsKey(i))
		{
			i++;
		}

		return i;
	}
	internal void SetHFLegacyDrawingRel(string relID) => SetXmlNodeString("d:legacyDrawingHF/@r:id", relID);
	internal void RemoveLegacyDrawingRel(string relID)
	{
		var n = WorksheetXml.DocumentElement.SelectSingleNode(string.Format("d:legacyDrawing[@r:id=\"{0}\"]", relID), NameSpaceManager);
		n?.ParentNode.RemoveChild(n);
	}

	internal void UpdateCellsWithDate1904Setting()
	{
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values);
		var offset = Workbook.Date1904 ? -ExcelWorkbook.date1904Offset : ExcelWorkbook.date1904Offset;
		while (cse.MoveNext())
		{
			if (cse.Value._value is DateTime time)
			{
				try
				{
					var sdv = time.ToOADate();
					sdv += offset;

					SetValueInner(cse.Row, cse.Column, DateTime.FromOADate(sdv));
				}
				catch
				{
					// A date outside the OLE Automation range is left unchanged.
				}
			}
		}
	}
	internal string GetFormula(int row, int col)
	{
		var v = _formulas.GetValue(row, col);
		if (v is int v1)
		{
			return _sharedFormulas[v1].GetFormula(row, col, Name);
		}
		else
		{
			return v != null ? v.ToString() : "";
		}
	}
	internal string GetFormulaR1C1(int row, int col)
	{
		var v = _formulas.GetValue(row, col);
		if (v is int v1)
		{
			var sf = _sharedFormulas[v1];
			return ExcelCellBase.TranslateToR1C1(sf.Formula, sf.StartRow, sf.StartCol);
		}
		else
		{
			return v != null ? ExcelCellBase.TranslateToR1C1(v.ToString(), row, col) : "";
		}
	}
}
