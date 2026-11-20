using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// The styleobject for the range.
	/// </summary>
	public ExcelStyle Style
	{
		get
		{
			IsRangeValid("styling");
			var s = 0;
			if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref s)) //Cell exists
			{
				if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref s)) //No, check Row style
				{
					var c = Worksheet.GetColumn(_fromCol);
					if (c == null)
					{
						s = 0;
					}
					else
					{
						s = c.StyleID;
					}
				}
			}

			return _worksheet.Workbook.Styles.GetStyleObject(s, _worksheet.PositionID, Address);
		}
	}
	private int GetColumnStyle(int col)
	{
		object c = null;
		if (_worksheet.ExistsValueInner(0, col, ref c))
		{
			return (c as ExcelColumn).StyleID;
		}
		else
		{
			var row = 0;
			if (_worksheet._values.PrevCell(ref row, ref col))
			{
				var column = _worksheet.GetValueInner(row, col) as ExcelColumn;
				if (column.ColumnMax >= col)
				{
					return _worksheet.GetStyleInner(row, col);
				}
			}
		}

		return 0;
	}
	/// <summary>
	/// The style ID. 
	/// It is not recomended to use this one. Use Named styles as an alternative.
	/// If you do, make sure that you use the Style.UpdateXml() method to update any new styles added to the workbook.
	/// </summary>
	public int StyleID
	{
		get
		{
			var s = 0;
			if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref s))
			{
				if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref s))
				{
					s = _worksheet.GetStyleInner(0, _fromCol);
				}
			}

			return s;
		}
		set
		{
			_changePropMethod(this, _setStyleIdDelegate, value);
		}
	}
}
