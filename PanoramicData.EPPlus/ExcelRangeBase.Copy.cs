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
		Dictionary<int, int> styleCashe = [];

		// Clear all existing cells
		int toRow = _toRow - _fromRow + 1,
			toCol = _toCol - _fromCol + 1;

		var i = 0;
		object o = null;
		byte flag = 0;
		Uri hl = null;

		var excludeFormulas = excelRangeCopyOptionFlags.HasValue && (excelRangeCopyOptionFlags.Value & ExcelRangeCopyOptionFlags.ExcludeFormulas) == ExcelRangeCopyOptionFlags.ExcludeFormulas;
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);

		var copiedValue = new List<CopiedCell>();
		while (cse.Next())
		{
			var row = cse.Row;
			var col = cse.Column;       //Issue 15070
			var cell = new CopiedCell
			{
				Row = Destination._fromRow + (row - _fromRow),
				Column = Destination._fromCol + (col - _fromCol),
				Value = cse.Value._value
			};

			if (!excludeFormulas && _worksheet._formulas.Exists(row, col, ref o))
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

			if (_worksheet.ExistsStyleInner(row, col, ref i))
			{
				if (sameWorkbook)
				{
					cell.StyleID = i;
				}
				else
				{
					if (styleCashe.TryGetValue(i, out var value))
					{
						i = value;
					}
					else
					{
						var oldStyleID = i;
						i = styles.CloneStyle(sourceStyles, i);
						styleCashe.Add(oldStyleID, i);
					}

					cell.StyleID = i;
				}
			}

			if (_worksheet._hyperLinks.Exists(row, col, ref hl))
			{
				cell.HyperLink = hl;
			}

			// Will just be null if no comment exists.
			cell.Comment = _worksheet.Cells[cse.Row, cse.Column].Comment;

			if (_worksheet._flags.Exists(row, col, ref flag))
			{
				cell.Flag = flag;
			}

			copiedValue.Add(cell);
		}

		//Copy styles with no cell value
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

				i = cses.Value._styleId;
				if (sameWorkbook)
				{
					cell.StyleID = i;
				}
				else
				{
					if (styleCashe.TryGetValue(i, out var value))
					{
						i = value;
					}
					else
					{
						var oldStyleID = i;
						i = styles.CloneStyle(sourceStyles, i);
						styleCashe.Add(oldStyleID, i);
					}

					cell.StyleID = i;
				}

				copiedValue.Add(cell);
			}
		}

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

		Destination._worksheet.MergedCells.Clear(new ExcelAddressBase(Destination._fromRow, Destination._fromCol, Destination._fromRow + toRow - 1, Destination._fromCol + toCol - 1));

		Destination._worksheet._values.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
		Destination._worksheet._formulas.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
		Destination._worksheet._hyperLinks.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
		Destination._worksheet._flags.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
		Destination._worksheet._commentsStore.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);

		foreach (var cell in copiedValue)
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

		//Add merged cells
		foreach (var m in copiedMergedCells.Values)
		{
			if (m != null)
			{
				Destination._worksheet.MergedCells.Add(m, true);
			}
		}

		//Check that the range is not larger than the dimensions of the worksheet. 
		//If so set the copied range to the worksheet dimensions to avoid copying empty cells.
		// Upstream behaviour kept: its collision test was always true, so any dimension is used.
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
