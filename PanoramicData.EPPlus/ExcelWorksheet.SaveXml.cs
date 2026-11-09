using System.Text;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void SaveXml(Stream stream)
	{
		//Create the nodes if they do not exist.
		StreamWriter sw = new(stream, Encoding.UTF8, 65536);
		if (this is ExcelChartsheet)
		{
			sw.Write(_worksheetXml.OuterXml);
		}
		else
		{
			CreateNode("d:cols");
			CreateNode("d:sheetData");
			CreateNode("d:mergeCells");
			CreateNode("d:hyperlinks");
			CreateNode("d:rowBreaks");
			CreateNode("d:colBreaks");

			var xml = _worksheetXml.OuterXml;
			int colStart = 0, colEnd = 0;
			GetBlockPos(xml, "cols", ref colStart, ref colEnd);

			sw.Write(xml[..colStart]);
			UpdateColumnData(sw);

			int cellStart = colEnd, cellEnd = colEnd;
			GetBlockPos(xml, "sheetData", ref cellStart, ref cellEnd);

			sw.Write(xml[colEnd..cellStart]);
			UpdateRowCellData(sw);

			int mergeStart = cellEnd, mergeEnd = cellEnd;

			GetBlockPos(xml, "mergeCells", ref mergeStart, ref mergeEnd);
			sw.Write(xml[cellEnd..mergeStart]);

			CleanupMergedCells(_mergedCells);
			if (_mergedCells.Count > 0)
			{
				UpdateMergedCells(sw);
			}

			int hyperStart = mergeEnd, hyperEnd = mergeEnd;
			GetBlockPos(xml, "hyperlinks", ref hyperStart, ref hyperEnd);
			sw.Write(xml[mergeEnd..hyperStart]);
			UpdateHyperLinks(sw);

			int rowBreakStart = hyperEnd, rowBreakEnd = hyperEnd;
			GetBlockPos(xml, "rowBreaks", ref rowBreakStart, ref rowBreakEnd);
			sw.Write(xml[hyperEnd..rowBreakStart]);
			UpdateRowBreaks(sw);

			int colBreakStart = rowBreakEnd, colBreakEnd = rowBreakEnd;
			GetBlockPos(xml, "colBreaks", ref colBreakStart, ref colBreakEnd);
			sw.Write(xml[rowBreakEnd..colBreakStart]);
			UpdateColBreaks(sw);
			sw.Write(xml[colBreakEnd..]);
		}

		sw.Flush();
	}

	private static void CleanupMergedCells(MergeCellsCollection _mergedCells)
	{
		var i = 0;
		while (i < _mergedCells.List.Count)
		{
			if (_mergedCells[i] == null)
			{
				_mergedCells.List.RemoveAt(i);
			}
			else
			{
				i++;
			}
		}
	}
	private void UpdateColBreaks(StreamWriter sw)
	{
		StringBuilder breaks = new();
		var count = 0;
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 0, 0, ExcelPackage.MaxColumns);
		while (cse.Next())
		{
			if (cse.Value._value is ExcelColumn col && col.PageBreak)
			{
				breaks.AppendFormat("<brk id=\"{0}\" max=\"16383\" man=\"1\"/>", cse.Column);
				count++;
			}
		}

		if (count > 0)
		{
			sw.Write(string.Format("<colBreaks count=\"{0}\" manualBreakCount=\"{0}\">{1}</colBreaks>", count, breaks));
		}
	}

	private void UpdateRowBreaks(StreamWriter sw)
	{
		StringBuilder breaks = new();
		var count = 0;
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 0, ExcelPackage.MaxRows, 0);
		while (cse.Next())
		{
			if (cse.Value._value is RowInternal row && row.PageBreak)
			{
				breaks.AppendFormat("<brk id=\"{0}\" max=\"1048575\" man=\"1\"/>", cse.Row);
				count++;
			}
		}

		if (count > 0)
		{
			sw.Write(string.Format("<rowBreaks count=\"{0}\" manualBreakCount=\"{0}\">{1}</rowBreaks>", count, breaks));
		}
	}
}
