using OfficeOpenXml.Compatibility;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property.
	/// Note: Cells containing formulas must be calculated before autofit is called.
	/// Wrapped and merged cells are also ignored.
	/// </summary>
	public void AutoFitColumns() => AutoFitColumns(_worksheet.DefaultColWidth);

	/// <summary>
	/// Set the column width from the content of the range.
	/// Note: Cells containing formulas are ignored if no calculation is made.
	///       Wrapped and merged cells are also ignored.
	/// </summary>
	/// <param name="MinimumWidth">Minimum column width</param>
	public void AutoFitColumns(double MinimumWidth) => AutoFitColumns(MinimumWidth, double.MaxValue);

	/// <summary>
	/// Set the column width from the content of the range.
	/// Note: Cells containing formulas are ignored if no calculation is made.
	///       Wrapped and merged cells are also ignored.
	///      Hidden columns are left hidden.
	/// </summary>
	/// <param name="MinimumWidth">Minimum column width</param>
	/// <param name="MaximumWidth">Maximum column width</param>
	public void AutoFitColumns(double MinimumWidth, double MaximumWidth)
	{
		if (_worksheet.Dimension == null)
		{
			return;
		}

		if (_fromCol < 1 || _fromRow < 1)
		{
			SetToSelectedRange();
		}


		var doAdjust = _worksheet._package.DoAdjustDrawings;
		_worksheet._package.DoAdjustDrawings = false;
		var drawWidths = _worksheet.Drawings.GetDrawingWidths();

		var fromCol = _fromCol > _worksheet.Dimension._fromCol ? _fromCol : _worksheet.Dimension._fromCol;
		var toCol = _toCol < _worksheet.Dimension._toCol ? _toCol : _worksheet.Dimension._toCol;

		if (fromCol > toCol)
		{
			return; //Issue 15383
		}

		if (Addresses == null)
		{
			SetMinWidth(MinimumWidth, fromCol, toCol);
		}
		else
		{
			foreach (var addr in Addresses)
			{
				fromCol = addr._fromCol > _worksheet.Dimension._fromCol ? addr._fromCol : _worksheet.Dimension._fromCol;
				toCol = addr._toCol < _worksheet.Dimension._toCol ? addr._toCol : _worksheet.Dimension._toCol;
				SetMinWidth(MinimumWidth, fromCol, toCol);
			}
		}

		//Get any auto filter to widen these columns
		var afAddr = new List<ExcelAddressBase>();
		if (_worksheet.AutoFilterAddress != null)
		{
			afAddr.Add(new ExcelAddressBase(_worksheet.AutoFilterAddress._fromRow,
												_worksheet.AutoFilterAddress._fromCol,
												_worksheet.AutoFilterAddress._fromRow,
												_worksheet.AutoFilterAddress._toCol));
			afAddr[^1]._ws = WorkSheet;
		}

		foreach (var tbl in _worksheet.Tables)
		{
			if (tbl.AutoFilterAddress != null)
			{
				afAddr.Add(new ExcelAddressBase(tbl.AutoFilterAddress._fromRow,
																		tbl.AutoFilterAddress._fromCol,
																		tbl.AutoFilterAddress._fromRow,
																		tbl.AutoFilterAddress._toCol));
				afAddr[^1]._ws = WorkSheet;
			}
		}

		var styles = _worksheet.Workbook.Styles;
		var nf = styles.Fonts[styles.CellXfs[0].FontId];
		var normalSize = Convert.ToSingle(ExcelWorkbook.GetWidthPixels(nf.Name, nf.Size));

		if (!TextMeasurer.IsAvailable)
		{
			//No usable font engine (native SkiaSharp library missing): leave the widths as they are.
			_worksheet.Drawings.AdjustWidth(drawWidths);
			_worksheet._package.DoAdjustDrawings = doAdjust;
			return;
		}

		using var measurer = new TextMeasurer();
		foreach (var cell in this)
		{
			if (_worksheet.Column(cell.Start.Column).Hidden)    //Issue 15338
			{
				continue;
			}

			if (cell.Merge == true || cell.Style.WrapText)
			{
				continue;
			}

			var fntID = styles.CellXfs[cell.StyleID].FontId;
			var fnt = styles.Fonts[fntID];
			var ind = styles.CellXfs[cell.StyleID].Indent;
			var textForWidth = cell.TextForWidth;
			var t = textForWidth + (ind > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', ind) : "");
			if (t.Length > 32000)
			{
				t = t[..32000]; //Issue
			}

			var size = measurer.Measure(t, fnt.Name, fnt.Size, fnt.Bold, fnt.Italic);

			double width;
			double r = styles.CellXfs[cell.StyleID].TextRotation;
			if (r <= 0)
			{
				width = (size.Width + 5) / normalSize;
			}
			else
			{
				r = (r <= 90 ? r : r - 90);
				width = (((size.Width - size.Height) * Math.Abs(Math.Cos(Math.PI * r / 180.0)) + size.Height) + 5) / normalSize;
			}

			foreach (var a in afAddr)
			{
				if (a.Collide(cell) != eAddressCollition.No)
				{
					width += 2.25;
					break;
				}
			}

			if (width > _worksheet.Column(cell._fromCol).Width)
			{
				_worksheet.Column(cell._fromCol).Width = width > MaximumWidth ? MaximumWidth : width;
			}
		}

		_worksheet.Drawings.AdjustWidth(drawWidths);
		_worksheet._package.DoAdjustDrawings = doAdjust;
	}

	private void SetMinWidth(double minimumWidth, int fromCol, int toCol)
	{
		var iterator = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 0, fromCol, 0, toCol);
		var prevCol = fromCol;
		foreach (var val in iterator)
		{
			var col = (ExcelColumn)val._value;
			col.Width = minimumWidth;
			if (_worksheet.DefaultColWidth > minimumWidth && col.ColumnMin > prevCol)
			{
				var newCol = _worksheet.Column(prevCol);
				newCol.ColumnMax = col.ColumnMin - 1;
				newCol.Width = minimumWidth;
			}

			prevCol = col.ColumnMax + 1;
		}

		if (_worksheet.DefaultColWidth > minimumWidth && prevCol < toCol)
		{
			var newCol = _worksheet.Column(prevCol);
			newCol.ColumnMax = toCol;
			newCol.Width = minimumWidth;
		}
	}

	internal string TextForWidth => GetFormattedText(true);
}
