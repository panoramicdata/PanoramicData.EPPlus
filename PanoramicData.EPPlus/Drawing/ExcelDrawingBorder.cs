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
 * Jan Källman		                Initial Release		        2009-12-22
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Type of Line cap
/// </summary>
public enum eLineCap
{
	Flat,   //flat
	Round,  //rnd
	Square  //Sq
}
/// <summary>
/// Line style.
/// </summary>
public enum eLineStyle
{
	Dash,
	DashDot,
	Dot,
	LongDash,
	LongDashDot,
	LongDashDotDot,
	Solid,
	SystemDash,
	SystemDashDot,
	SystemDashDotDot,
	SystemDot
}
/// <summary>
/// Border for drawings
/// </summary>    
public sealed class ExcelDrawingBorder : XmlHelper
{
	string _linePath;
	internal ExcelDrawingBorder(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath) :
		base(nameSpaceManager, topNode)
	{
		SchemaNodeOrder = ["chart", "tickLblPos", "spPr", "txPr", "crossAx", "printSettings", "showVal", "showCatName", "showSerName", "showPercent", "separator", "showLeaderLines", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill", "prstDash"];
		_linePath = linePath;
		_lineStylePath = string.Format(_lineStylePath, linePath);
		_lineCapPath = string.Format(_lineCapPath, linePath);
		_lineWidth = string.Format(_lineWidth, linePath);
	}
	#region "Public properties"
	ExcelDrawingFill _fill = null;
	/// <summary>
	/// Fill
	/// </summary>
	public ExcelDrawingFill Fill
	{
		get
		{
			_fill ??= new ExcelDrawingFill(NameSpaceManager, TopNode, _linePath);

			return _fill;
		}
	}
	string _lineStylePath = "{0}/a:prstDash/@val";
	/// <summary>
	/// Linestyle
	/// </summary>
	public eLineStyle LineStyle
	{
		get
		{
			return TranslateLineStyle(GetXmlNodeString(_lineStylePath));
		}
		set
		{
			CreateNode(_linePath, false);
			SetXmlNodeString(_lineStylePath, TranslateLineStyleText(value));
		}
	}
	string _lineCapPath = "{0}/@cap";
	/// <summary>
	/// Linecap
	/// </summary>
	public eLineCap LineCap
	{
		get
		{
			return TranslateLineCap(GetXmlNodeString(_lineCapPath));
		}
		set
		{
			CreateNode(_linePath, false);
			SetXmlNodeString(_lineCapPath, TranslateLineCapText(value));
		}
	}
	string _lineWidth = "{0}/@w";
	/// <summary>
	/// Width in pixels
	/// </summary>
	public int Width
	{
		get
		{
			return GetXmlNodeInt(_lineWidth) / 12700;
		}
		set
		{
			SetXmlNodeString(_lineWidth, (value * 12700).ToString());
		}
	}
	#endregion
	#region "Translate Enum functions"
	private string TranslateLineStyleText(eLineStyle value)
	{
		var text = value.ToString();
		return value switch
		{
			eLineStyle.Dash or eLineStyle.Dot or eLineStyle.DashDot or eLineStyle.Solid => text[..1].ToLower(CultureInfo.InvariantCulture) + text[1..],//First to Lower case.
			eLineStyle.LongDash or eLineStyle.LongDashDot or eLineStyle.LongDashDotDot => "lg" + text[4..],
			eLineStyle.SystemDash or eLineStyle.SystemDashDot or eLineStyle.SystemDashDotDot or eLineStyle.SystemDot => "sys" + text[6..],
			_ => throw (new Exception("Invalid Linestyle")),
		};
	}
	private eLineStyle TranslateLineStyle(string text) => text switch
	{
		"dash" or "dot" or "dashDot" or "solid" => (eLineStyle)Enum.Parse(typeof(eLineStyle), text, true),
		"lgDash" or "lgDashDot" or "lgDashDotDot" => (eLineStyle)Enum.Parse(typeof(eLineStyle), "Long" + text[2..]),
		"sysDash" or "sysDashDot" or "sysDashDotDot" or "sysDot" => (eLineStyle)Enum.Parse(typeof(eLineStyle), "System" + text[3..]),
		_ => throw (new Exception("Invalid Linestyle")),
	};
	private string TranslateLineCapText(eLineCap value) => value switch
	{
		eLineCap.Round => "rnd",
		eLineCap.Square => "sq",
		_ => "flat",
	};
	private eLineCap TranslateLineCap(string text) => text switch
	{
		"rnd" => eLineCap.Round,
		"sq" => eLineCap.Square,
		_ => eLineCap.Flat,
	};
	#endregion


	//public ExcelDrawingFont Font
	//{
	//    get
	//    { 

	//    }
	//}
}
