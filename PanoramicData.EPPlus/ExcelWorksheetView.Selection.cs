using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheetView
{
	private XmlElement SelectionNode
	{
		get
		{
			if (SheetViewElement.SelectSingleNode("//d:selection", _worksheet.NameSpaceManager) is not XmlElement selectionNode)
			{
				selectionNode = _worksheet.WorksheetXml.CreateElement("selection", ExcelPackage.schemaMain);
				SheetViewElement.AppendChild(selectionNode);
			}

			return selectionNode;
		}
	}
	/// <summary>
	/// The active cell. Single Cell address.                
	/// This cell must be inside the selected range. If not, the selected range is set to the active cell address
	/// </summary>
	public string ActiveCell
	{
		get
		{
			return Panes[Panes.GetUpperBound(0)].ActiveCell;
		}
		set
		{
			var ac = new ExcelAddressBase(value);
			if (ac.IsSingleCell == false)
			{
				throw new InvalidOperationException("ActiveCell must be a single cell.");
			}

			/*** Active cell must be inside SelectedRange ***/
			var sd = new ExcelAddressBase(SelectedRange.Replace(" ", ","));
			Panes[Panes.GetUpperBound(0)].ActiveCell = value;

			if (IsActiveCellInSelection(ac, sd) == false)
			{
				SelectedRange = value;
			}
		}
	}

	/// <summary>
	/// Selected Cells in the worksheet. Used in combination with ActiveCell.
	/// If the active cell is not inside the selected range, the active cell will be set to the first cell in the selected range.
	/// If the selected range has multiple adresses, these are separated with space. If the active cell is not within the first address in this list, the attribute ActiveCellId must be set (not supported, so it must be set via the XML).
	/// </summary>
	public string SelectedRange
	{
		get
		{
			return Panes[Panes.GetUpperBound(0)].SelectedRange;
		}
		set
		{
			var ac = new ExcelAddressBase(ActiveCell);

			/*** Active cell must be inside SelectedRange ***/
			var sd = new ExcelAddressBase(value.Replace(" ", ","));      //Space delimitered here, replace

			Panes[Panes.GetUpperBound(0)].SelectedRange = value;
			if (IsActiveCellInSelection(ac, sd) == false)
			{
				ActiveCell = new ExcelCellAddress(sd._fromRow, sd._fromCol).Address;
			}
		}
	}

	private static bool IsActiveCellInSelection(ExcelAddressBase ac, ExcelAddressBase sd)
	{
		var c = sd.Collide(ac);
		if (c is ExcelAddressBase.eAddressCollition.Equal or ExcelAddressBase.eAddressCollition.Inside)
		{
			return true;
		}
		else
		{
			if (sd.Addresses != null)
			{
				foreach (var sds in sd.Addresses)
				{
					c = sds.Collide(ac);
					if (c is ExcelAddressBase.eAddressCollition.Equal or ExcelAddressBase.eAddressCollition.Inside)
					{
						return true;
					}
				}
			}
		}

		return false;
	}
}
