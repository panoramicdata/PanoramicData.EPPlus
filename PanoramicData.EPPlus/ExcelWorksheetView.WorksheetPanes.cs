using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheetView
{
	/// <summary>
	/// The worksheet panes after a freeze or split.
	/// </summary>
	public class ExcelWorksheetPanes : XmlHelper
	{
		XmlElement _selectionNode;
		internal ExcelWorksheetPanes(XmlNamespaceManager ns, XmlNode topNode) :
			base(ns, topNode)
		{
			if (topNode.Name == "selection")
			{
				_selectionNode = topNode as XmlElement;
			}
		}

		const string _activeCellPath = "@activeCell";
		/// <summary>
		/// Set the active cell. Must be set within the SelectedRange.
		/// </summary>
		public string ActiveCell
		{
			get
			{
				var address = GetXmlNodeString(_activeCellPath);
				return address == "" ? "A1" : address;
			}
			set
			{
				if (_selectionNode == null)
				{
					CreateSelectionElement();
				}

				ExcelCellBase.GetRowColFromAddress(value, out var fromRow, out var fromCol, out int toRow, out var toCol);
				SetXmlNodeString(_activeCellPath, value);
				if (((XmlElement)TopNode).GetAttribute("sqref") == "")
				{

					SelectedRange = ExcelCellBase.GetAddress(fromRow, fromCol);
				}
				else
				{
					//TODO:Add fix for out of range here
				}
			}
		}

		private void CreateSelectionElement()
		{
			_selectionNode = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
			TopNode.AppendChild(_selectionNode);
			TopNode = _selectionNode;
		}
		const string _selectionRangePath = "@sqref";
		/// <summary>
		/// Selected Cells.Used in combination with ActiveCell
		/// </summary>        
		public string SelectedRange
		{
			get
			{
				var address = GetXmlNodeString(_selectionRangePath);
				return address == "" ? "A1" : address;
			}
			set
			{
				if (_selectionNode == null)
				{
					CreateSelectionElement();
				}

				ExcelCellBase.GetRowColFromAddress(value, out var fromRow, out var fromCol, out int toRow, out var toCol);
				SetXmlNodeString(_selectionRangePath, value);
				if (((XmlElement)TopNode).GetAttribute("activeCell") == "")
				{

					ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);
				}
				else
				{
					//TODO:Add fix for out of range here
				}
			}
		}
	}
}
