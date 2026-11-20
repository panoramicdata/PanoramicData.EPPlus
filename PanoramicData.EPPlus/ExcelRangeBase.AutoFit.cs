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

		if (!SetMinWidths(MinimumWidth))
		{
			return; //Issue 15383
		}

		//Get any auto filter to widen these columns
		var afAddr = GetAutoFilterAddresses();

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
			AutoFitCell(cell, measurer, styles, normalSize, afAddr, MaximumWidth);
		}

		_worksheet.Drawings.AdjustWidth(drawWidths);
		_worksheet._package.DoAdjustDrawings = doAdjust;
	}

	private bool SetMinWidths(double MinimumWidth)
	{
		var fromCol = Math.Max(_fromCol, _worksheet.Dimension._fromCol);
		var toCol = Math.Min(_toCol, _worksheet.Dimension._toCol);

		if (fromCol > toCol)
		{
			return false;
		}

		if (Addresses == null)
		{
			SetMinWidth(MinimumWidth, fromCol, toCol);
		}
		else
		{
			foreach (var addr in Addresses)
			{
				fromCol = Math.Max(addr._fromCol, _worksheet.Dimension._fromCol);
				toCol = Math.Min(addr._toCol, _worksheet.Dimension._toCol);
				SetMinWidth(MinimumWidth, fromCol, toCol);
			}
		}

		return true;
	}

	private List<ExcelAddressBase> GetAutoFilterAddresses()
	{
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

		return afAddr;
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
