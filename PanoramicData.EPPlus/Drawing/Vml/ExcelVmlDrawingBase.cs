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
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Horizontal Alingment
/// </summary>
public enum eTextAlignHorizontalVml
{
	Left,
	Center,
	Right
}
/// <summary>
/// Vertical Alingment
/// </summary>
public enum eTextAlignVerticalVml
{
	Top,
	Center,
	Bottom
}
/// <summary>
/// Linestyle
/// </summary>
public enum eLineStyleVml
{
	Solid,
	Round,
	Square,
	Dash,
	DashDot,
	LongDash,
	LongDashDot,
	LongDashDotDot
}
/// <summary>
/// Drawing object used for comments
/// </summary>
public class ExcelVmlDrawingBase : XmlHelper
{
	internal ExcelVmlDrawingBase(XmlNode topNode, XmlNamespaceManager ns) :
		base(ns, topNode)
	{
		SchemaNodeOrder = ["fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked", "AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible"];
	}
	public string Id
	{
		get
		{
			return GetXmlNodeString("@id");
		}
		set
		{
			SetXmlNodeString("@id", value);
		}
	}
	/// <summary>
	/// Alternative text to be displayed instead of a graphic.
	/// </summary>
	public string AlternativeText
	{
		get
		{
			return GetXmlNodeString("@alt");
		}
		set
		{
			SetXmlNodeString("@alt", value);
		}
	}
	#region "Style Handling methods"
	protected static bool GetStyle(string style, string key, out string value)
	{
		var styles = style.Split(';');
		foreach (var s in styles)
		{
			if (s.IndexOf(':') > 0)
			{
				var split = s.Split(':');
				if (split[0] == key)
				{
					value = split[1];
					return true;
				}
			}
			else if (s == key)
			{
				value = "";
				return true;
			}
		}

		value = "";
		return false;
	}
	protected static string SetStyle(string style, string key, string value)
	{
		var styles = style.Split([';'], StringSplitOptions.RemoveEmptyEntries);
		var newStyle = "";
		var changed = false;
		foreach (var s in styles)
		{
			var split = s.Split(':');
			if (split[0].Trim() == key)
			{
				if (value.Trim() != "") //If blank remove the item
				{
					newStyle += key + ':' + value;
				}

				changed = true;
			}
			else
			{
				newStyle += s;
			}

			newStyle += ';';
		}

		if (!changed)
		{
			newStyle += key + ':' + value;
		}
		else
		{
			newStyle = newStyle[..^1];
		}

		return newStyle;
	}
	#endregion
}
