using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheetView
{

	private ExcelWorksheetPanes[] LoadPanes()
	{
		var nodes = TopNode.SelectNodes("//d:selection", NameSpaceManager);
		if (nodes.Count == 0)
		{
			return [new(NameSpaceManager, TopNode)];
		}
		else
		{
			var panes = new ExcelWorksheetPanes[nodes.Count];
			var i = 0;
			foreach (XmlElement elem in nodes)
			{
				panes[i++] = new ExcelWorksheetPanes(NameSpaceManager, elem);
			}

			return panes;
		}
	}
	/// <summary>
	/// Reference to the panes
	/// </summary>
	public ExcelWorksheetPanes[] Panes
	{
		get;
		internal set;
	}
	/// <summary>
	/// Freeze the columns/rows to left and above the cell
	/// </summary>
	/// <param name="Row"></param>
	/// <param name="Column"></param>
	public void FreezePanes(int Row, int Column)
	{
		//TODO:fix this method to handle splits as well.
		if (Row == 1 && Column == 1)
		{
			UnFreezePanes();
		}

		string sqRef = SelectedRange, activeCell = ActiveCell;

		var paneNode = CreateFrozenPaneNode(Row, Column);

		RemoveSelection();

		if (Row > 1 && Column == 1)
		{
			paneNode.SetAttribute("activePane", "bottomLeft");
			var sel = CreatePaneSelection("bottomLeft", activeCell, sqRef);
			sel.SetAttribute("sqref", sqRef);
			TopNode.InsertAfter(sel, paneNode);
		}
		else if (Column > 1 && Row == 1)
		{
			paneNode.SetAttribute("activePane", "topRight");
			var sel = CreatePaneSelection("topRight", activeCell, sqRef);
			TopNode.InsertAfter(sel, paneNode);
		}
		else
		{
			AddBottomRightPaneSelections(paneNode, Row, Column, activeCell, sqRef);
		}

		Panes = LoadPanes();
	}

	private XmlElement CreateFrozenPaneNode(int Row, int Column)
	{
		var paneNode = TopNode.SelectSingleNode(_paneNodePath, NameSpaceManager) as XmlElement;
		if (paneNode == null)
		{
			CreateNode(_paneNodePath);
			paneNode = TopNode.SelectSingleNode(_paneNodePath, NameSpaceManager) as XmlElement;
		}

		paneNode.RemoveAll();   //Clear all attributes
		if (Column > 1)
		{
			paneNode.SetAttribute("xSplit", (Column - 1).ToString());
		}

		if (Row > 1)
		{
			paneNode.SetAttribute("ySplit", (Row - 1).ToString());
		}

		paneNode.SetAttribute("topLeftCell", ExcelCellBase.GetAddress(Row, Column));
		paneNode.SetAttribute("state", "frozen");
		return paneNode;
	}

	private XmlElement CreatePaneSelection(string pane, string activeCell, string sqRef)
	{
		var sel = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
		sel.SetAttribute("pane", pane);
		if (activeCell != "")
		{
			sel.SetAttribute("activeCell", activeCell);
		}

		if (sqRef != "")
		{
			sel.SetAttribute("sqref", sqRef);
		}

		return sel;
	}

	private void AddBottomRightPaneSelections(XmlElement paneNode, int Row, int Column, string activeCell, string sqRef)
	{
		paneNode.SetAttribute("activePane", "bottomRight");
		var sel1 = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
		sel1.SetAttribute("pane", "topRight");
		var cell = ExcelCellBase.GetAddress(1, Column);
		sel1.SetAttribute("activeCell", cell);
		sel1.SetAttribute("sqref", cell);
		paneNode.ParentNode.InsertAfter(sel1, paneNode);

		var sel2 = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
		cell = ExcelCellBase.GetAddress(Row, 1);
		sel2.SetAttribute("pane", "bottomLeft");
		sel2.SetAttribute("activeCell", cell);
		sel2.SetAttribute("sqref", cell);
		sel1.ParentNode.InsertAfter(sel2, sel1);

		var sel3 = CreatePaneSelection("bottomRight", activeCell, sqRef);
		sel2.ParentNode.InsertAfter(sel3, sel2);
	}

	private void RemoveSelection()
	{
		//Find selection nodes and remove them            
		var selections = TopNode.SelectNodes(_selectionNodePath, NameSpaceManager);
		foreach (XmlNode sel in selections)
		{
			sel.ParentNode.RemoveChild(sel);
		}
	}
	/// <summary>
	/// Unlock all rows and columns to scroll freely
	/// /// </summary>
	public void UnFreezePanes()
	{
		string sqRef = SelectedRange, activeCell = ActiveCell;

		var paneNode = TopNode.SelectSingleNode(_paneNodePath, NameSpaceManager) as XmlElement;
		paneNode?.ParentNode.RemoveChild(paneNode);

		RemoveSelection();

		Panes = LoadPanes();

		SelectedRange = sqRef;
		ActiveCell = activeCell;
	}
}
