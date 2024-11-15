﻿/*******************************************************************************
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
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Globalization;
using System.Xml;
using System.Drawing;

namespace OfficeOpenXml.Style;

/// <summary>
/// Linestyle
/// </summary>
public enum eUnderLineType
{
	Dash,
	DashHeavy,
	DashLong,
	DashLongHeavy,
	Double,
	DotDash,
	DotDashHeavy,
	DotDotDash,
	DotDotDashHeavy,
	Dotted,
	DottedHeavy,
	Heavy,
	None,
	Single,
	Wavy,
	WavyDbl,
	WavyHeavy,
	Words
}
/// <summary>
/// Type of font strike
/// </summary>
public enum eStrikeType
{
	Double,
	No,
	Single
}
/// <summary>
/// Used by Rich-text and Paragraphs.
/// </summary>
public class ExcelTextFont : XmlHelper
{
	readonly string _path;
	readonly XmlNode _rootNode;
	internal ExcelTextFont(XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder)
		: base(namespaceManager, rootNode)
	{
		SchemaNodeOrder = schemaNodeOrder;
		_rootNode = rootNode;
		if (path != "")
		{
			var node = rootNode.SelectSingleNode(path, namespaceManager);
			if (node != null)
			{
				TopNode = node;
			}
		}

		_path = path;
	}
	readonly string _fontLatinPath = "a:latin/@typeface";
	public string LatinFont
	{
		get
		{
			return GetXmlNodeString(_fontLatinPath);
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_fontLatinPath, value);
		}
	}

	protected internal void CreateTopNode()
	{
		if (_path != "" && TopNode == _rootNode)
		{
			CreateNode(_path);
			TopNode = _rootNode.SelectSingleNode(_path, NameSpaceManager);
		}
	}
	readonly string _fontCsPath = "a:cs/@typeface";
	public string ComplexFont
	{
		get
		{
			return GetXmlNodeString(_fontCsPath);
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_fontCsPath, value);
		}
	}
	readonly string _boldPath = "@b";
	public bool Bold
	{
		get
		{
			return GetXmlNodeBool(_boldPath);
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_boldPath, value ? "1" : "0");
		}
	}
	readonly string _underLinePath = "@u";
	public eUnderLineType UnderLine
	{
		get
		{
			return TranslateUnderline(GetXmlNodeString(_underLinePath));
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_underLinePath, TranslateUnderlineText(value));
		}
	}
	readonly string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";
	public Color UnderLineColor
	{
		get
		{
			var col = GetXmlNodeString(_underLineColorPath);
			return col == "" ? Color.Empty : Color.FromArgb(int.Parse(col, NumberStyles.AllowHexSpecifier));
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_underLineColorPath, value.ToArgb().ToString("X").Substring(2, 6));
		}
	}
	readonly string _italicPath = "@i";
	public bool Italic
	{
		get
		{
			return GetXmlNodeBool(_italicPath);
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_italicPath, value ? "1" : "0");
		}
	}
	readonly string _strikePath = "@strike";
	public eStrikeType Strike
	{
		get
		{
			return TranslateStrike(GetXmlNodeString(_strikePath));
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_strikePath, TranslateStrikeText(value));
		}
	}
	readonly string _sizePath = "@sz";
	public float Size
	{
		get
		{
			return GetXmlNodeInt(_sizePath) / 100;
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_sizePath, ((int)(value * 100)).ToString());
		}
	}
	readonly string _colorPath = "a:solidFill/a:srgbClr/@val";
	public Color Color
	{
		get
		{
			var col = GetXmlNodeString(_colorPath);
			return col == "" ? Color.Empty : Color.FromArgb(int.Parse(col, NumberStyles.AllowHexSpecifier));
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_colorPath, value.ToArgb().ToString("X").Substring(2, 6));
		}
	}
	#region "Translate methods"
	private static eUnderLineType TranslateUnderline(string text) => text switch
	{
		"sng" => eUnderLineType.Single,
		"dbl" => eUnderLineType.Double,
		"" => eUnderLineType.None,
		_ => Enum.Parse<eUnderLineType>(text),
	};
	private static string TranslateUnderlineText(eUnderLineType value)
	{
		switch (value)
		{
			case eUnderLineType.Single:
				return "sng";
			case eUnderLineType.Double:
				return "dbl";
			default:
				var ret = value.ToString();
				return ret[..1].ToLower(CultureInfo.InvariantCulture) + ret[1..];
		}
	}
	private static eStrikeType TranslateStrike(string text) => text switch
	{
		"dblStrike" => eStrikeType.Double,
		"sngStrike" => eStrikeType.Single,
		_ => eStrikeType.No,
	};
	private static string TranslateStrikeText(eStrikeType value) => value switch
	{
		eStrikeType.Single => "sngStrike",
		eStrikeType.Double => "dblStrike",
		_ => "noStrike",
	};
	#endregion
	/// <summary>
	/// Set the font style from a font object
	/// </summary>
	/// <param name="Font"></param>
	public void SetFromFont(Font Font)
	{
		LatinFont = Font.Name;
		ComplexFont = Font.Name;
		Size = Font.Size;
		if (Font.Bold) Bold = Font.Bold;
		if (Font.Italic) Italic = Font.Italic;
		if (Font.Underline) UnderLine = eUnderLineType.Single;
		if (Font.Strikeout) Strike = eStrikeType.Single;
	}
}
