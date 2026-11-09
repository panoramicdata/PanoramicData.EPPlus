namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Copies the range of cells to an other range
	/// </summary>
	/// <param name="Destination">The start cell where the range will be copied.</param>
	public void Copy(ExcelRangeBase Destination) => Copy(Destination, null);

	/// <summary>
	/// Copies the range of cells to an other range
	/// </summary>
	/// <param name="Destination">The start cell where the range will be copied.</param>
	/// <param name="excelRangeCopyOptionFlags">Cell parts that will not be copied. If Formulas are specified, the formulas will NOT be copied.</param>
	public void Copy(ExcelRangeBase Destination, ExcelRangeCopyOptionFlags? excelRangeCopyOptionFlags)
	{
		var sameWorkbook = Destination._worksheet.Workbook == _worksheet.Workbook;
		ExcelStyles sourceStyles = _worksheet.Workbook.Styles,
					styles = Destination._worksheet.Workbook.Styles;
		var styleMap = new CopyStyleMap(sameWorkbook, sourceStyles, styles);

		// Clear all existing cells
		int toRow = _toRow - _fromRow + 1,
			toCol = _toCol - _fromCol + 1;

		var excludeFormulas = excelRangeCopyOptionFlags.HasValue && (excelRangeCopyOptionFlags.Value & ExcelRangeCopyOptionFlags.ExcludeFormulas) == ExcelRangeCopyOptionFlags.ExcludeFormulas;

		var copiedValue = GetCopiedCells(Destination, excludeFormulas, styleMap);

		//Copy styles with no cell value
		AddCopiedStylesWithoutValue(copiedValue, Destination, styleMap);

		var copiedMergedCells = GetCopiedMergedCells(Destination);

		ClearCopyDestination(Destination, toRow, toCol);

		foreach (var cell in copiedValue)
		{
			WriteCopiedCell(Destination, cell);
		}

		//Add merged cells
		foreach (var m in copiedMergedCells.Values)
		{
			if (m != null)
			{
				Destination._worksheet.MergedCells.Add(m, true);
			}
		}

		CopyOutlineLevels(Destination);
	}

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

	private Dictionary<int, ExcelAddress?> GetCopiedMergedCells(ExcelRangeBase Destination)
	{
		var copiedMergedCells = new Dictionary<int, ExcelAddress?>();

		//Merged cells
		var csem = new CellsStoreEnumerator<int>(_worksheet.MergedCells._cells, _fromRow, _fromCol, _toRow, _toCol);
		while (csem.Next())
		{
			if (!copiedMergedCells.ContainsKey(csem.Value))
			{
				var adr = new ExcelAddress(_worksheet.Name, _worksheet.MergedCells.List[csem.Value]);
				var collideResult = Collide(adr);
				if (collideResult is eAddressCollition.Inside or eAddressCollition.Equal)
				{
					copiedMergedCells.Add(csem.Value, new ExcelAddress(
						Destination._fromRow + (adr.Start.Row - _fromRow),
						Destination._fromCol + (adr.Start.Column - _fromCol),
						Destination._fromRow + (adr.End.Row - _fromRow),
						Destination._fromCol + (adr.End.Column - _fromCol)));
				}
				else
				{
					//Partial merge of the address ignore.
					copiedMergedCells.Add(csem.Value, null);
				}
			}
		}

		return copiedMergedCells;
	}

	private static void ClearCopyDestination(ExcelRangeBase Destination, int toRow, int toCol)
	{
		Destination._worksheet.MergedCells.Clear(new ExcelAddressBase(Destination._fromRow, Destination._fromCol, Destination._fromRow + toRow - 1, Destination._fromCol + toCol - 1));

		Destination._worksheet._values.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
		Destination._worksheet._formulas.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
		Destination._worksheet._hyperLinks.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
		Destination._worksheet._flags.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
		Destination._worksheet._commentsStore.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
	}

	private void WriteCopiedCell(ExcelRangeBase Destination, CopiedCell cell)
	{
		Destination._worksheet.SetValueInner(cell.Row, cell.Column, cell.Value);

		if (cell.StyleID != null)
		{
			Destination._worksheet.SetStyleInner(cell.Row, cell.Column, cell.StyleID.Value);
		}

		if (cell.Formula != null)
		{
			cell.Formula = UpdateFormulaReferences(cell.Formula.ToString(), Destination._fromRow - _fromRow, Destination._fromCol - _fromCol, 0, 0, Destination.WorkSheet, Destination.WorkSheet, true);
			Destination._worksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
		}

		if (cell.HyperLink != null)
		{
			Destination._worksheet._hyperLinks.SetValue(cell.Row, cell.Column, cell.HyperLink);
		}

		if (cell.Comment != null)
		{
			Destination.Worksheet.Cells[cell.Row, cell.Column].AddComment(cell.Comment.Text, cell.Comment.Author);
		}

		if (cell.Flag != 0)
		{
			Destination._worksheet._flags.SetValue(cell.Row, cell.Column, cell.Flag);
		}
	}

	private void CopyOutlineLevels(ExcelRangeBase Destination)
	{
		// Upstream behaviour kept: the worksheet dimension, when there is one, bounds the outline levels copied.
		ExcelAddressBase range = Worksheet.Dimension ?? (ExcelAddressBase)this;

		if (_fromCol == 1 && _toCol == ExcelPackage.MaxColumns)
		{
			for (var r = 0; r < range.Rows; r++)
			{
				var destinationRow = Destination.Worksheet.Row(Destination.Start.Row + r);
				destinationRow.OutlineLevel = Worksheet.Row(range._fromRow + r).OutlineLevel;
			}
		}

		if (_fromRow == 1 && _toRow == ExcelPackage.MaxRows)
		{
			for (var c = 0; c < range.Columns; c++)
			{
				var destinationCol = Destination.Worksheet.Column(Destination.Start.Column + c);
				destinationCol.OutlineLevel = Worksheet.Column(range._fromCol + c).OutlineLevel;
			}
		}
	}
	/// <summary>
	/// Clear all cells
	/// </summary>
	public void Clear() => Delete(this, false);
}
