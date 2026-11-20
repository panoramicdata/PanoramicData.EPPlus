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
 * Jan Källman		    Initial Release		        2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;
using System.Diagnostics.CodeAnalysis;
using System.Xml;
namespace OfficeOpenXml;

/// <summary>
/// Containts all shared cell styles for a workbook
/// </summary>
[SuppressMessage("Minor Code Smell", "S1104:Fields should not have public accessibility", Justification = "Public API: the style collections are published fields; converting them to properties would break binary compatibility.")]
[SuppressMessage("Minor Code Smell", "S2357:Fields should be private", Justification = "Public API: the style collections are published fields; converting them to properties would break binary compatibility.")]
public sealed partial class ExcelStyles : XmlHelper
{
	const string NumberFormatsPath = "d:styleSheet/d:numFmts";
	const string FontsPath = "d:styleSheet/d:fonts";
	const string FillsPath = "d:styleSheet/d:fills";
	const string BordersPath = "d:styleSheet/d:borders";
	const string CellStyleXfsPath = "d:styleSheet/d:cellStyleXfs";
	const string CellXfsPath = "d:styleSheet/d:cellXfs";
	const string CellStylesPath = "d:styleSheet/d:cellStyles";
	const string dxfsPath = "d:styleSheet/d:dxfs";

	readonly XmlDocument _styleXml;
	readonly ExcelWorkbook _wb;
	readonly XmlNamespaceManager _nameSpaceManager;
	internal int _nextDfxNumFmtID = 164;
	internal ExcelStyles(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb) :
		base(NameSpaceManager, xml)
	{
		_styleXml = xml;
		_wb = wb;
		_nameSpaceManager = NameSpaceManager;
		SchemaNodeOrder = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs"];
		LoadFromDocument();
	}
	/// <summary>
	/// Loads the style XML to memory
	/// </summary>
	private void LoadFromDocument()
	{
		LoadNumberFormatsAndFonts();
		LoadFillsAndBorders();
		LoadXfs();
		LoadNamedStylesAndDxfs();
	}

	private void LoadNumberFormatsAndFonts()
	{
		//NumberFormats
		ExcelNumberFormatXml.AddBuildIn(NameSpaceManager, NumberFormats);
		var numNode = _styleXml.SelectSingleNode(NumberFormatsPath, _nameSpaceManager);
		if (numNode != null)
		{
			foreach (XmlNode n in numNode)
			{
				ExcelNumberFormatXml nf = new(_nameSpaceManager, n);
				NumberFormats.Add(nf.Id, nf);
				if (nf.NumFmtId >= NumberFormats.NextId)
				{
					NumberFormats.NextId = nf.NumFmtId + 1;
				}
			}
		}

		//Fonts
		var fontNode = _styleXml.SelectSingleNode(FontsPath, _nameSpaceManager);
		foreach (XmlNode n in fontNode)
		{
			ExcelFontXml f = new(_nameSpaceManager, n);
			Fonts.Add(f.Id, f);
		}
	}

	private void LoadFillsAndBorders()
	{
		//Fills
		var fillNode = _styleXml.SelectSingleNode(FillsPath, _nameSpaceManager);
		foreach (XmlNode n in fillNode)
		{
			ExcelFillXml f;
			if (n.FirstChild != null && n.FirstChild.LocalName == "gradientFill")
			{
				f = new ExcelGradientFillXml(_nameSpaceManager, n);
			}
			else
			{
				f = new ExcelFillXml(_nameSpaceManager, n);
			}

			Fills.Add(f.Id, f);
		}

		//Borders
		var borderNode = _styleXml.SelectSingleNode(BordersPath, _nameSpaceManager);
		foreach (XmlNode n in borderNode)
		{
			ExcelBorderXml b = new(_nameSpaceManager, n);
			Borders.Add(b.Id, b);
		}
	}

	private void LoadXfs()
	{
		//cellStyleXfs
		var styleXfsNode = _styleXml.SelectSingleNode(CellStyleXfsPath, _nameSpaceManager);
		if (styleXfsNode != null)
		{
			foreach (XmlNode n in styleXfsNode)
			{
				ExcelXfs item = new(_nameSpaceManager, n, this);
				CellStyleXfs.Add(item.Id, item);
			}
		}

		var styleNode = _styleXml.SelectSingleNode(CellXfsPath, _nameSpaceManager);
		for (var i = 0; i < styleNode.ChildNodes.Count; i++)
		{
			var n = styleNode.ChildNodes[i];
			ExcelXfs item = new(_nameSpaceManager, n, this);
			CellXfs.Add(item.Id, item);
		}
	}

	private void LoadNamedStylesAndDxfs()
	{
		//cellStyle
		var namedStyleNode = _styleXml.SelectSingleNode(CellStylesPath, _nameSpaceManager);
		if (namedStyleNode != null)
		{
			foreach (XmlNode n in namedStyleNode)
			{
				ExcelNamedStyleXml item = new(_nameSpaceManager, n, this);
				NamedStyles.Add(item.Name, item);
			}
		}

		//dxfsPath
		var dxfsNode = _styleXml.SelectSingleNode(dxfsPath, _nameSpaceManager);
		if (dxfsNode != null)
		{
			foreach (XmlNode x in dxfsNode)
			{
				ExcelDxfStyleConditionalFormatting item = new(_nameSpaceManager, x, this);
				Dxfs.Add(item.Id, item);
			}
		}
	}
	internal ExcelStyle GetStyleObject(int Id, int PositionID, string Address)
	{
		if (Id < 0)
		{
			Id = 0;
		}

		return new ExcelStyle(this, PropertyChange, PositionID, Address, Id);
	}
	public ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats = [];
	public ExcelStyleCollection<ExcelFontXml> Fonts = [];
	public ExcelStyleCollection<ExcelFillXml> Fills = [];
	public ExcelStyleCollection<ExcelBorderXml> Borders = [];
	public ExcelStyleCollection<ExcelXfs> CellStyleXfs = [];
	public ExcelStyleCollection<ExcelXfs> CellXfs = [];
	public ExcelStyleCollection<ExcelNamedStyleXml> NamedStyles = [];
	public ExcelStyleCollection<ExcelDxfStyleConditionalFormatting> Dxfs = [];

	internal static string Id => "";
	private static int GetXmlNodeInt(XmlNode node) => int.TryParse(GetXmlNode(node), out var i) ? i : 0;
	private static string GetXmlNode(XmlNode node)
	{
		if (node == null)
		{
			return "";
		}

		return node.Value != null ? node.Value : "";
	}
}
