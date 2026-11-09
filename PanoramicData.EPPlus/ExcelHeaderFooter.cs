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
 *******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman                      Total rewrite               2010-03-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 * *******************************************************************************/
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils;
using System.Diagnostics.CodeAnalysis;
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// Represents the Header and Footer on an Excel Worksheet
/// </summary>
[SuppressMessage("Critical Code Smell", "S2339:Public constant members should not be used", Justification = "Public API: the format codes are published constants used in constant expressions by callers.")]
public sealed partial class ExcelHeaderFooter : XmlHelper
{
	/// <summary>
	/// The code for "current page #"
	/// </summary>
	public const string PageNumber = @"&P";
	/// <summary>
	/// The code for "total pages"
	/// </summary>
	public const string NumberOfPages = @"&N";
	/// <summary>
	/// The code for "text font color"
	/// RGB Color is specified as RRGGBB
	/// Theme Color is specified as TTSNN where TT is the theme color Id, S is either "+" or "-" of the tint/shade value, NN is the tint/shade value.
	/// </summary>
	public const string FontColor = @"&K";
	/// <summary>
	/// The code for "sheet tab name"
	/// </summary>
	public const string SheetName = @"&A";
	/// <summary>
	/// The code for "this workbook's file path"
	/// </summary>
	public const string FilePath = @"&Z";
	/// <summary>
	/// The code for "this workbook's file name"
	/// </summary>
	public const string FileName = @"&F";
	/// <summary>
	/// The code for "date"
	/// </summary>
	public const string CurrentDate = @"&D";
	/// <summary>
	/// The code for "time"
	/// </summary>
	public const string CurrentTime = @"&T";
	/// <summary>
	/// The code for "picture as background"
	/// </summary>
	public const string Image = @"&G";
	/// <summary>
	/// The code for "outline style"
	/// </summary>
	public const string OutlineStyle = @"&O";
	/// <summary>
	/// The code for "shadow style"
	/// </summary>
	public const string ShadowStyle = @"&H";

	internal ExcelHeaderFooterText _oddHeader;
	internal ExcelHeaderFooterText _oddFooter;
	internal ExcelHeaderFooterText _evenHeader;
	internal ExcelHeaderFooterText _evenFooter;
	internal ExcelHeaderFooterText _firstHeader;
	internal ExcelHeaderFooterText _firstFooter;
	private readonly ExcelWorksheet _ws;

	/// <summary>
	/// ExcelHeaderFooter Constructor
	/// </summary>
	/// <param name="nameSpaceManager"></param>
	/// <param name="topNode"></param>
	/// <param name="ws">The worksheet</param>
	internal ExcelHeaderFooter(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelWorksheet ws) :
		base(nameSpaceManager, topNode)
	{
		_ws = ws;
		SchemaNodeOrder = ["headerFooter", "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"];
	}

	const string alignWithMarginsPath = "@alignWithMargins";
	/// <summary>
	/// Gets/sets the alignWithMargins attribute
	/// </summary>
	public bool AlignWithMargins
	{
		get
		{
			return GetXmlNodeBool(alignWithMarginsPath);
		}
		set
		{
			SetXmlNodeString(alignWithMarginsPath, value ? "1" : "0");
		}
	}

	const string differentOddEvenPath = "@differentOddEven";
	/// <summary>
	/// Gets/sets the flag that tells Excel to display different headers and footers on odd and even pages.
	/// </summary>
	public bool differentOddEven
	{
		get
		{
			return GetXmlNodeBool(differentOddEvenPath);
		}
		set
		{
			SetXmlNodeString(differentOddEvenPath, value ? "1" : "0");
		}
	}

	const string differentFirstPath = "@differentFirst";

	/// <summary>
	/// Gets/sets the flag that tells Excel to display different headers and footers on the first page of the worksheet.
	/// </summary>
	public bool differentFirst
	{
		get
		{
			return GetXmlNodeBool(differentFirstPath);
		}
		set
		{
			SetXmlNodeString(differentFirstPath, value ? "1" : "0");
		}
	}
	const string scaleWithDocPath = "@scaleWithDoc";
	/// <summary>
	/// Specify whether the header and footer should scale as you use the "Shrink to fit" feature on the document
	/// </summary>
	public bool ScaleWithDocument
	{
		get
		{
			return GetXmlNodeBool(scaleWithDocPath);
		}
		set
		{
			SetXmlNodeBool(scaleWithDocPath, value);
		}
	}
	/// <summary>
	/// Provides access to the header on odd numbered pages of the document.
	/// If you want the same header on both odd and even pages, then only set values in this ExcelHeaderFooterText class.
	/// </summary>
	public ExcelHeaderFooterText OddHeader
	{
		get
		{
			_oddHeader ??= new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:oddHeader", NameSpaceManager), _ws, "H");

			return _oddHeader;
		}
	}
	/// <summary>
	/// Provides access to the footer on odd numbered pages of the document.
	/// If you want the same footer on both odd and even pages, then only set values in this ExcelHeaderFooterText class.
	/// </summary>
	public ExcelHeaderFooterText OddFooter
	{
		get
		{
			if (_oddFooter == null)
			{
				_oddFooter = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:oddFooter", NameSpaceManager), _ws, "F");
			}

			return _oddFooter;
		}
	}
	// evenHeader and evenFooter set differentOddEven = true
	/// <summary>
	/// Provides access to the header on even numbered pages of the document.
	/// </summary>
	public ExcelHeaderFooterText EvenHeader
	{
		get
		{
			if (_evenHeader == null)
			{
				_evenHeader = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:evenHeader", NameSpaceManager), _ws, "HEVEN");
				differentOddEven = true;
			}

			return _evenHeader;
		}
	}
	/// <summary>
	/// Provides access to the footer on even numbered pages of the document.
	/// </summary>
	public ExcelHeaderFooterText EvenFooter
	{
		get
		{
			if (_evenFooter == null)
			{
				_evenFooter = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:evenFooter", NameSpaceManager), _ws, "FEVEN");
				differentOddEven = true;
			}

			return _evenFooter;
		}
	}
	/// <summary>
	/// Provides access to the header on the first page of the document.
	/// </summary>
	public ExcelHeaderFooterText FirstHeader
	{
		get
		{
			if (_firstHeader == null)
			{
				_firstHeader = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:firstHeader", NameSpaceManager), _ws, "HFIRST");
				differentFirst = true;
			}

			return _firstHeader;
		}
	}
	/// <summary>
	/// Provides access to the footer on the first page of the document.
	/// </summary>
	public ExcelHeaderFooterText FirstFooter
	{
		get
		{
			if (_firstFooter == null)
			{
				_firstFooter = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:firstFooter", NameSpaceManager), _ws, "FFIRST");
				differentFirst = true;
			}

			return _firstFooter;
		}
	}
	private ExcelVmlDrawingPictureCollection _vmlDrawingsHF;
	/// <summary>
	/// Vml drawings. Underlaying object for Header footer images
	/// </summary>
	public ExcelVmlDrawingPictureCollection Pictures
	{
		get
		{
			if (_vmlDrawingsHF == null)
			{
				var vmlNode = _ws.WorksheetXml.SelectSingleNode("d:worksheet/d:legacyDrawingHF/@r:id", NameSpaceManager);
				if (vmlNode == null)
				{
					_vmlDrawingsHF = new ExcelVmlDrawingPictureCollection(_ws._package, _ws, null);
				}
				else
				{
					if (_ws.Part.RelationshipExists(vmlNode.Value))
					{
						var rel = _ws.Part.GetRelationship(vmlNode.Value);
						var vmlUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

						_vmlDrawingsHF = new ExcelVmlDrawingPictureCollection(_ws._package, _ws, vmlUri)
						{
							RelId = rel.Id
						};
					}
				}
			}

			return _vmlDrawingsHF;
		}
	}
}
