using OfficeOpenXml.Compatibility;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private void AutoFitCell(ExcelRangeBase cell, TextMeasurer measurer, ExcelStyles styles, float normalSize, List<ExcelAddressBase> afAddr, double MaximumWidth)
	{
		if (IsSkippedByAutoFit(cell))
		{
			return;
		}

		var width = MeasureCellWidth(cell, measurer, styles, normalSize, afAddr);
		if (width > _worksheet.Column(cell._fromCol).Width)
		{
			_worksheet.Column(cell._fromCol).Width = width > MaximumWidth ? MaximumWidth : width;
		}
	}

	//Issue 15338: hidden columns are left hidden
	private bool IsSkippedByAutoFit(ExcelRangeBase cell)
		=> _worksheet.Column(cell.Start.Column).Hidden || cell.Merge == true || cell.Style.WrapText;

	private static double MeasureCellWidth(ExcelRangeBase cell, TextMeasurer measurer, ExcelStyles styles, float normalSize, List<ExcelAddressBase> afAddr)
	{
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

		if (CollidesWithAny(afAddr, cell))
		{
			width += 2.25;
		}

		return width;
	}

	private static bool CollidesWithAny(List<ExcelAddressBase> addresses, ExcelRangeBase cell)
	{
		foreach (var a in addresses)
		{
			if (a.Collide(cell) != eAddressCollition.No)
			{
				return true;
			}
		}

		return false;
	}
}
