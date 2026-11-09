namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private sealed class CopyStyleMap(bool sameWorkbook, ExcelStyles sourceStyles, ExcelStyles styles)
	{
		private readonly Dictionary<int, int> _styleCashe = [];

		internal int Map(int styleId)
		{
			if (sameWorkbook)
			{
				return styleId;
			}

			if (_styleCashe.TryGetValue(styleId, out var value))
			{
				return value;
			}

			var newStyleId = styles.CloneStyle(sourceStyles, styleId);
			_styleCashe.Add(styleId, newStyleId);
			return newStyleId;
		}
	}

	private List<CopiedCell> GetCopiedCells(ExcelRangeBase Destination, bool excludeFormulas, CopyStyleMap styleMap)
	{
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);

		var copiedValue = new List<CopiedCell>();
		while (cse.Next())
		{
			copiedValue.Add(CreateCopiedCell(cse, Destination, excludeFormulas, styleMap));
		}

		return copiedValue;
	}

	private CopiedCell CreateCopiedCell(CellsStoreEnumerator<ExcelCoreValue> cse, ExcelRangeBase Destination, bool excludeFormulas, CopyStyleMap styleMap)
	{
		var row = cse.Row;
		var col = cse.Column;       //Issue 15070
		var cell = new CopiedCell
		{
			Row = Destination._fromRow + (row - _fromRow),
			Column = Destination._fromCol + (col - _fromCol),
			Value = cse.Value._value
		};

		if (!excludeFormulas)
		{
			CopyCellFormula(cse, Destination, cell);
		}

		var i = 0;
		if (_worksheet.ExistsStyleInner(row, col, ref i))
		{
			cell.StyleID = styleMap.Map(i);
		}

		Uri hl = null;
		if (_worksheet._hyperLinks.Exists(row, col, ref hl))
		{
			cell.HyperLink = hl;
		}

		// Will just be null if no comment exists.
		cell.Comment = _worksheet.Cells[cse.Row, cse.Column].Comment;

		byte flag = 0;
		if (_worksheet._flags.Exists(row, col, ref flag))
		{
			cell.Flag = flag;
		}

		return cell;
	}

	private void CopyCellFormula(CellsStoreEnumerator<ExcelCoreValue> cse, ExcelRangeBase Destination, CopiedCell cell)
	{
		object o = null;
		if (_worksheet._formulas.Exists(cse.Row, cse.Column, ref o))
		{
			if (o is int)
			{
				cell.Formula = _worksheet.GetFormula(cse.Row, cse.Column);
				if (_worksheet._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.ArrayFormula))
				{
					Destination._worksheet._flags.SetFlagValue(cse.Row, cse.Column, true, CellFlags.ArrayFormula);
				}
			}
			else
			{
				cell.Formula = o;
			}
		}
	}

	private void AddCopiedStylesWithoutValue(List<CopiedCell> copiedValue, ExcelRangeBase Destination, CopyStyleMap styleMap)
	{
		var cses = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
		while (cses.Next())
		{
			if (!_worksheet.ExistsValueInner(cses.Row, cses.Column))
			{
				var row = Destination._fromRow + (cses.Row - _fromRow);
				var col = Destination._fromCol + (cses.Column - _fromCol);
				var cell = new CopiedCell
				{
					Row = row,
					Column = col,
					Value = null
				};

				cell.StyleID = styleMap.Map(cses.Value._styleId);

				copiedValue.Add(cell);
			}
		}
	}
}
