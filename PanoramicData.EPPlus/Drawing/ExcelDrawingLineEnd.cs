﻿using System;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Line end style.
/// </summary>
public enum eEndStyle   //ST_LineEndType
{
	/// <summary>
	/// No end
	/// </summary>
	None,
	/// <summary>
	/// Triangle arrow head
	/// </summary>
	Triangle,
	/// <summary>
	/// Stealth arrow head
	/// </summary>
	Stealth,
	/// <summary>
	/// Diamond
	/// </summary>
	Diamond,
	/// <summary>
	/// Oval
	/// </summary>
	Oval,
	/// <summary>
	/// Line arrow head
	/// </summary>
	Arrow
}

/// <summary>
/// Lend end size.
/// </summary>
public enum eEndSize
{
	/// <summary>
	/// Smal
	/// </summary>
	Small,
	/// <summary>
	/// Medium
	/// </summary>
	Medium,
	/// <summary>
	/// Large
	/// </summary>
	Large
}

/// <summary>
/// Properties for drawing line ends
/// </summary>
public sealed class ExcelDrawingLineEnd : XmlHelper
{
	string _linePath;
	internal ExcelDrawingLineEnd(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath) :
		base(nameSpaceManager, topNode)
	{
		SchemaNodeOrder = ["headEnd", "tailEnd"];
		_linePath = linePath;
	}
	string _headEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@type";
	/// <summary>
	/// HeaderEnd
	/// </summary>
	public eEndStyle HeadEnd
	{
		get
		{
			return TranslateEndStyle(GetXmlNodeString(_headEndStylePath));
		}
		set
		{
			CreateNode(_linePath, false);
			SetXmlNodeString(_headEndStylePath, TranslateEndStyleText(value));
		}
	}
	string _tailEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@type";
	/// <summary>
	/// HeaderEnd
	/// </summary>
	public eEndStyle TailEnd
	{
		get
		{
			return TranslateEndStyle(GetXmlNodeString(_tailEndStylePath));
		}
		set
		{
			CreateNode(_linePath, false);
			SetXmlNodeString(_tailEndStylePath, TranslateEndStyleText(value));
		}
	}

	string _tailEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@w";
	/// <summary>
	/// TailEndSizeWidth
	/// </summary>
	public eEndSize TailEndSizeWidth
	{
		get
		{
			return TranslateEndSize(GetXmlNodeString(_tailEndSizeWidthPath));
		}
		set
		{
			CreateNode(_linePath, false);
			SetXmlNodeString(_tailEndSizeWidthPath, TranslateEndSizeText(value));
		}
	}

	string _tailEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@len";
	/// <summary>
	/// TailEndSizeHeight
	/// </summary>
	public eEndSize TailEndSizeHeight
	{
		get
		{
			return TranslateEndSize(GetXmlNodeString(_tailEndSizeHeightPath));
		}
		set
		{
			CreateNode(_linePath, false);
			SetXmlNodeString(_tailEndSizeHeightPath, TranslateEndSizeText(value));
		}
	}

	string _headEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@w";
	/// <summary>
	/// TailEndSizeWidth
	/// </summary>
	public eEndSize HeadEndSizeWidth
	{
		get
		{
			return TranslateEndSize(GetXmlNodeString(_headEndSizeWidthPath));
		}
		set
		{
			CreateNode(_linePath, false);
			SetXmlNodeString(_headEndSizeWidthPath, TranslateEndSizeText(value));
		}
	}

	string _headEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@len";
	/// <summary>
	/// TailEndSizeHeight
	/// </summary>
	public eEndSize HeadEndSizeHeight
	{
		get
		{
			return TranslateEndSize(GetXmlNodeString(_headEndSizeHeightPath));
		}
		set
		{
			CreateNode(_linePath, false);
			SetXmlNodeString(_headEndSizeHeightPath, TranslateEndSizeText(value));
		}
	}

	#region "Translate Enum functions"
	private string TranslateEndStyleText(eEndStyle value) => value.ToString().ToLower();
	private eEndStyle TranslateEndStyle(string text) => text switch
	{
		"none" or "arrow" or "diamond" or "oval" or "stealth" or "triangle" => (eEndStyle)Enum.Parse(typeof(eEndStyle), text, true),
		_ => throw (new Exception("Invalid Endstyle")),
	};

	private string TranslateEndSizeText(eEndSize value)
	{
		var text = value.ToString();
		return value switch
		{
			eEndSize.Small => "sm",
			eEndSize.Medium => "med",
			eEndSize.Large => "lg",
			_ => throw (new Exception("Invalid Endsize")),
		};
	}
	private eEndSize TranslateEndSize(string text) => text switch
	{
		"sm" or "med" or "lg" => (eEndSize)Enum.Parse(typeof(eEndSize), text, true),
		_ => throw (new Exception("Invalid Endsize")),
	};
	#endregion
}