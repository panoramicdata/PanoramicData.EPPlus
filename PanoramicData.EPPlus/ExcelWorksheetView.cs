/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		    Initial Release		       2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System.Diagnostics.CodeAnalysis;
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// Represents the different view states of the worksheet
/// </summary>
public partial class ExcelWorksheetView : XmlHelper
{
	private readonly ExcelWorksheet _worksheet;

	/// <summary>
	/// Creates a new ExcelWorksheetView which provides access to all the view states of the worksheet.
	/// </summary>
	/// <param name="ns"></param>
	/// <param name="node"></param>
	/// <param name="xlWorksheet"></param>
	internal ExcelWorksheetView(XmlNamespaceManager ns, XmlNode node, ExcelWorksheet xlWorksheet) :
		base(ns, node)
	{
		_worksheet = xlWorksheet;
		SchemaNodeOrder = ["sheetViews", "sheetView", "pane", "selection"];
		Panes = LoadPanes();
	}
	/// <summary>
	/// Returns a reference to the sheetView element
	/// </summary>
	protected internal XmlElement SheetViewElement => (XmlElement)TopNode;

	/// <summary>
	/// Indicates if the worksheet is selected within the workbook. NOTE: Setter clears other selected tabs.
	/// </summary>
	public bool TabSelected
	{
		get
		{
			return GetXmlNodeBool("@tabSelected");
		}
		set
		{
			SetTabSelected(value);
		}
	}

	/// <summary>
	/// Indicates if the worksheet is selected within the workbook. NOTE: Setter keeps other selected tabs.
	/// </summary>
	public bool TabSelectedMulti
	{
		get
		{
			return GetXmlNodeBool("@tabSelected");
		}
		set
		{
			SetTabSelected(value, true);
		}
	}

	/// <summary>
	/// Sets whether the worksheet is selected within the workbook.
	/// </summary>
	/// <param name="isSelected">Whether the tab is selected, defaults to true.</param>
	/// <param name="allowMultiple">Whether to allow multiple active tabs, defaults to false.</param>
	[SuppressMessage("Major Code Smell", "S2360:Optional parameters should not be used", Justification = "Public API: replacing the optional parameters with overloads would break binary compatibility.")]
	public void SetTabSelected(bool isSelected = true, bool allowMultiple = false)
	{
		if (isSelected)
		{
			SheetViewElement.SetAttribute("tabSelected", "1");
			if (!allowMultiple)
			{
				// ensure no other worksheet has its tabSelected attribute set to 1
				foreach (var sheet in _worksheet._package.Workbook.Worksheets)
				{
					sheet.View.TabSelected = false;
				}
			}

			var bookView = _worksheet.Workbook.WorkbookXml.SelectSingleNode("//d:workbookView", _worksheet.NameSpaceManager) as XmlElement;
			bookView?.SetAttribute("activeTab", (_worksheet.PositionID - 1).ToString());
		}
		else
		{
			SetXmlNodeString("@tabSelected", "0");
		}
	}

	/// <summary>
	/// Sets the view mode of the worksheet to pagelayout
	/// </summary>
	public bool PageLayoutView
	{
		get
		{
			return GetXmlNodeString("@view") == "pageLayout";
		}
		set
		{
			if (value)
			{
				SetXmlNodeString("@view", "pageLayout");
			}
			else
			{
				SheetViewElement.RemoveAttribute("view");
			}
		}
	}
	/// <summary>
	/// Sets the view mode of the worksheet to pagebreak
	/// </summary>
	public bool PageBreakView
	{
		get
		{
			return GetXmlNodeString("@view") == "pageBreakPreview";
		}
		set
		{
			if (value)
			{
				SetXmlNodeString("@view", "pageBreakPreview");
			}
			else
			{
				SheetViewElement.RemoveAttribute("view");
			}
		}
	}
	/// <summary>
	/// Show gridlines in the worksheet
	/// </summary>
	public bool ShowGridLines
	{
		get
		{
			return GetXmlNodeBool("@showGridLines", true);
		}
		set
		{
			SetXmlNodeString("@showGridLines", value ? "1" : "0");
		}
	}
	/// <summary>
	/// Show the Column/Row headers (containg column letters and row numbers)
	/// </summary>
	public bool ShowHeaders
	{
		get
		{
			return GetXmlNodeBool("@showRowColHeaders");
		}
		set
		{
			SetXmlNodeString("@showRowColHeaders", value ? "1" : "0");
		}
	}
	/// <summary>
	/// Window zoom magnification for current view representing percent values.
	/// </summary>
	public int ZoomScale
	{
		get
		{
			return GetXmlNodeInt("@zoomScale");
		}
		set
		{
			if (value is < 10 or > 400)
			{
				throw new ArgumentOutOfRangeException("Zoome scale out of range (10-400)");
			}

			SetXmlNodeString("@zoomScale", value.ToString());
		}
	}
	/// <summary>
	/// Flag indicating whether the sheet is in 'right to left' display mode. When in this mode,Column A is on the far right, Column B ;is one column left of Column A, and so on. Also,information in cells is displayed in the Right to Left format.
	/// </summary>
	public bool RightToLeft
	{
		get
		{
			return GetXmlNodeBool("@rightToLeft");
		}
		set
		{
			SetXmlNodeString("@rightToLeft", value == true ? "1" : "0");
		}
	}
	internal bool WindowProtection
	{
		get
		{
			return GetXmlNodeBool("@windowProtection", false);
		}
		set
		{
			SetXmlNodeBool("@windowProtection", value, false);
		}
	}
	readonly string _paneNodePath = "d:pane";
	readonly string _selectionNodePath = "d:selection";
}
