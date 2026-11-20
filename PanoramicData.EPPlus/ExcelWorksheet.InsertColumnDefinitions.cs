namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void ShiftColumnDefinitions(int columnFrom, int columns)
	{
		var lst = GetColumnDefinitions();

		for (var i = lst.Count - 1; i >= 0; i--)
		{
			ShiftColumnDefinition(lst[i], columnFrom, columns);
		}
	}

	private List<ExcelColumn> GetColumnDefinitions()
	{
		var csec = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 1, 0, ExcelPackage.MaxColumns);
		var lst = new List<ExcelColumn>();
		foreach (var val in csec)
		{
			var col = val._value;
			if (col is ExcelColumn column)
			{
				lst.Add(column);
			}
		}

		return lst;
	}

	private void ShiftColumnDefinition(ExcelColumn c, int columnFrom, int columns)
	{
		if (c._columnMin >= columnFrom)
		{
			if (c._columnMin + columns <= ExcelPackage.MaxColumns)
			{
				c._columnMin += columns;
			}
			else
			{
				c._columnMin = ExcelPackage.MaxColumns;
			}

			if (c._columnMax + columns <= ExcelPackage.MaxColumns)
			{
				c._columnMax += columns;
			}
			else
			{
				c._columnMax = ExcelPackage.MaxColumns;
			}
		}
		else if (c._columnMax >= columnFrom)
		{
			var cc = c._columnMax - columnFrom;
			c._columnMax = columnFrom - 1;
			CopyColumn(c, columnFrom + columns, columnFrom + columns + cc);
		}
	}

	private void CopyColumnStyles(int columnFrom, int columns, int copyStylesFromColumn)
	{
		if (copyStylesFromColumn >= columnFrom)
		{
			copyStylesFromColumn += columns;
		}

		//Get styles to a cached list, 
		var l = GetColumnStyles(copyStylesFromColumn);

		//Set the style id's from the list.
		foreach (var sc in l)
		{
			for (var c = 0; c < columns; c++)
			{
				if (sc[0] == 0)
				{
					var col = Column(columnFrom + c);   //Create the column
					col.StyleID = sc[1];
				}
				else
				{
					SetStyleInner(sc[0], columnFrom + c, sc[1]);
				}
			}
		}

		var newOutlineLevel = Column(copyStylesFromColumn).OutlineLevel;
		for (var c = 0; c < columns; c++)
		{
			Column(columnFrom + c).OutlineLevel = newOutlineLevel;
		}
	}

	private List<int[]> GetColumnStyles(int copyStylesFromColumn)
	{
		var l = new List<int[]>();
		var sce = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, copyStylesFromColumn, ExcelPackage.MaxRows, copyStylesFromColumn);
		lock (sce)
		{
			while (sce.Next())
			{
				if (sce.Value._styleId == 0)
				{
					continue;
				}

				l.Add([sce.Row, sce.Value._styleId]);
			}
		}

		return l;
	}
}