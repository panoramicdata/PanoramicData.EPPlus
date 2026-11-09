using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml;

public sealed partial class ExcelStyles
{
	internal int CloneStyle(ExcelStyles style, int styleID) => CloneStyle(style, styleID, false, false);
	internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle) => CloneStyle(style, styleID, isNamedStyle, false);
	internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle, bool allwaysAddCellXfs)
	{
		ExcelXfs xfs;
		lock (style)
		{
			if (isNamedStyle)
			{
				xfs = style.CellStyleXfs[styleID];
			}
			else
			{
				xfs = style.CellXfs[styleID];
			}

			var newXfs = xfs.Copy(this);
			//Numberformat
			if (xfs.NumberFormatId > 0)
			{
				CloneNumberFormat(style, xfs, newXfs);
			}

			CloneFontBorderFill(style, xfs, newXfs);

			//Named style reference
			if (xfs.XfId > 0)
			{
				CloneNamedStyleReference(style, xfs, newXfs, allwaysAddCellXfs);
			}

			return AddClonedXfs(newXfs, isNamedStyle, allwaysAddCellXfs);
		}
	}

	private void CloneNumberFormat(ExcelStyles style, ExcelXfs xfs, ExcelXfs newXfs)
	{
		//rake36: Two problems here...
		//rake36:  1. the first time through when format stays equal to String.Empty, it adds a string.empty to the list of Number Formats
		//rake36:  2. when adding a second sheet, if the numberformatid == 164, it finds the 164 added by previous sheets but was using the array index
		//rake36:      for the numberformatid

		var format = string.Empty;
		foreach (var fmt in style.NumberFormats)
		{
			if (fmt.NumFmtId == xfs.NumberFormatId)
			{
				format = fmt.Format;
				break;
			}
		}
		//rake36: Don't add another format if it's blank
		if (!String.IsNullOrEmpty(format))
		{
			var ix = NumberFormats.FindIndexByID(format);
			if (ix < 0)
			{
				var item = new ExcelNumberFormatXml(NameSpaceManager) { Format = format, NumFmtId = NumberFormats.NextId++ };
				NumberFormats.Add(format, item);
				//rake36: Use the just added format id
				newXfs.NumberFormatId = item.NumFmtId;
			}
			else
			{
				//rake36: Use the format id defined by the index... not the index itself
				newXfs.NumberFormatId = NumberFormats[ix].NumFmtId;
			}
		}
	}

	private void CloneFontBorderFill(ExcelStyles style, ExcelXfs xfs, ExcelXfs newXfs)
	{
		//Font
		if (xfs.FontId > -1)
		{
			var ix = Fonts.FindIndexByID(xfs.Font.Id);
			if (ix < 0)
			{
				var item = style.Fonts[xfs.FontId].Copy();
				ix = Fonts.Add(xfs.Font.Id, item);
			}

			newXfs.FontId = ix;
		}

		//Border
		if (xfs.BorderId > -1)
		{
			var ix = Borders.FindIndexByID(xfs.Border.Id);
			if (ix < 0)
			{
				var item = style.Borders[xfs.BorderId].Copy();
				ix = Borders.Add(xfs.Border.Id, item);
			}

			newXfs.BorderId = ix;
		}

		//Fill
		if (xfs.FillId > -1)
		{
			var ix = Fills.FindIndexByID(xfs.Fill.Id);
			if (ix < 0)
			{
				var item = style.Fills[xfs.FillId].Copy();
				ix = Fills.Add(xfs.Fill.Id, item);
			}

			newXfs.FillId = ix;
		}
	}

	private void CloneNamedStyleReference(ExcelStyles style, ExcelXfs xfs, ExcelXfs newXfs, bool allwaysAddCellXfs)
	{
		var id = style.CellStyleXfs[xfs.XfId].Id;
		var newId = CellStyleXfs.FindIndexByID(id);
		if (newId >= 0)
		{
			newXfs.XfId = newId;
		}
		else if (style._wb != _wb && allwaysAddCellXfs == false) //Not the same workbook, copy the namedstyle to the workbook or match the id
		{
			var nsFind = style.NamedStyles.ToDictionary(d => (d.StyleXfId));
			if (nsFind.TryGetValue(xfs.XfId, out var st))
			{
				if (NamedStyles.ExistsKey(st.Name))
				{
					newXfs.XfId = NamedStyles.FindIndexByID(st.Name);
				}
				else
				{
					CreateNamedStyle(st.Name, st.Style);
					newXfs.XfId = NamedStyles.Count - 1;
				}
			}
		}
	}

	private int AddClonedXfs(ExcelXfs newXfs, bool isNamedStyle, bool allwaysAddCellXfs)
	{
		int index;
		if (isNamedStyle && allwaysAddCellXfs == false)
		{
			index = CellStyleXfs.Add(newXfs.Id, newXfs);
		}
		else
		{
			if (allwaysAddCellXfs)
			{
				index = CellXfs.Add(newXfs.Id, newXfs);
			}
			else
			{
				index = CellXfs.FindIndexByID(newXfs.Id);
				if (index < 0)
				{
					index = CellXfs.Add(newXfs.Id, newXfs);
				}
			}
		}

		return index;
	}
}