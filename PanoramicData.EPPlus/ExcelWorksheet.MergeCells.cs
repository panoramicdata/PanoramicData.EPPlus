using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Load merged cells
	/// </summary>
	/// <param name="xr"></param>
	private void LoadMergeCells(XmlReader xr)
	{
		if (ReadUntil(xr, "mergeCells", "hyperlinks", "rowBreaks", "colBreaks") && !xr.EOF)
		{
			while (xr.Read())
			{
				if (xr.LocalName != "mergeCell")
				{
					break;
				}

				if (xr.NodeType == XmlNodeType.Element)
				{
					var address = xr.GetAttribute("ref");
					_mergedCells.Add(new ExcelAddress(address), false);
				}
			}
		}
	}

	/// <summary>
	/// Update merged cells
	/// </summary>
	/// <param name="sw">The writer</param>
	private void UpdateMergedCells(StreamWriter sw)
	{
		sw.Write("<mergeCells>");
		foreach (var address in _mergedCells)
		{
			sw.Write("<mergeCell ref=\"{0}\" />", address);
		}

		sw.Write("</mergeCells>");
	}

	/// <summary>
	/// Get MergeCell Index No
	/// </summary>
	/// <param name="row"></param>
	/// <param name="column"></param>
	/// <returns></returns>
	public int GetMergeCellId(int row, int column)
	{
		for (var i = 0; i < _mergedCells.Count; i++)
		{
			if (!string.IsNullOrEmpty(_mergedCells[i]))
			{
				var range = Cells[_mergedCells[i]];

				if (range.Start.Row <= row && row <= range.End.Row)
				{
					if (range.Start.Column <= column && column <= range.End.Column)
					{
						return i + 1;
					}
				}
			}
		}

		return 0;
	}
}
