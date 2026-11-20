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
 * Jan Källman		Added		21-MAR-2011
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// An Excel Pivottable
/// </summary>
public partial class ExcelPivotTable : XmlHelper
{
	internal ExcelPivotTable(Packaging.ZipPackageRelationship rel, ExcelWorksheet sheet) :
		base(sheet.NameSpaceManager)
	{
		WorkSheet = sheet;
		PivotTableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
		Relationship = rel;
		var pck = sheet._package.Package;
		Part = pck.GetPart(PivotTableUri);

		PivotTableXml = new XmlDocument();
		LoadXmlSafe(PivotTableXml, Part.GetStream());
		init();
		TopNode = PivotTableXml.DocumentElement;
		Address = new ExcelAddressBase(GetXmlNodeString("d:location/@ref"));

		_cacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this);
		LoadFields();

		LoadRowColumnFields("d:rowFields/d:field", RowFields);
		LoadRowColumnFields("d:colFields/d:field", ColumnFields);
		LoadPageAndDataFields();
	}

	private void LoadRowColumnFields(string path, ExcelPivotTableRowColumnFieldCollection collection)
	{
		foreach (XmlElement elem in TopNode.SelectNodes(path, NameSpaceManager))
		{
			if (int.TryParse(elem.GetAttribute("x"), out var x) && x >= 0)
			{
				collection.AddInternal(Fields[x]);
			}
			else
			{
				elem.ParentNode.RemoveChild(elem);
			}
		}
	}

	private void LoadPageAndDataFields()
	{
		foreach (XmlElement pageElem in TopNode.SelectNodes("d:pageFields/d:pageField", NameSpaceManager))
		{
			if (int.TryParse(pageElem.GetAttribute("fld"), out var fld) && fld >= 0)
			{
				var field = Fields[fld];
				field._pageFieldSettings = new ExcelPivotTablePageFieldSettings(NameSpaceManager, pageElem, field, fld);
				PageFields.AddInternal(field);
			}
		}

		foreach(XmlElement dataElem in TopNode.SelectNodes("d:dataFields/d:dataField", NameSpaceManager))
		{
			if (int.TryParse(dataElem.GetAttribute("fld"), out var fld) && fld >= 0)
			{
				var field = Fields[fld];
				var dataField = new ExcelPivotTableDataField(NameSpaceManager, dataElem, field);
				DataFields.AddInternal(dataField);
			}
		}
	}
	/// <summary>
	/// Add a new pivottable
	/// </summary>
	/// <param name="sheet">The worksheet</param>
	/// <param name="address">the address of the pivottable</param>
	/// <param name="sourceAddress">The address of the Source data</param>
	/// <param name="name"></param>
	/// <param name="tblId"></param>
	internal ExcelPivotTable(ExcelWorksheet sheet, ExcelAddressBase address, ExcelRangeBase sourceAddress, string name, int tblId) :
		base(sheet.NameSpaceManager)
	{
		WorkSheet = sheet;
		Address = address;
		var pck = sheet._package.Package;

		PivotTableXml = new XmlDocument();
		LoadXmlSafe(PivotTableXml, GetStartXml(name, tblId, address, sourceAddress), Encoding.UTF8);
		TopNode = PivotTableXml.DocumentElement;
		PivotTableUri = GetNewUri(pck, "/xl/pivotTables/pivotTable{0}.xml", ref tblId);
		init();

		Part = pck.CreatePart(PivotTableUri, ExcelPackage.schemaPivotTable);
		PivotTableXml.Save(Part.GetStream());

		//Worksheet-Pivottable relationship
		Relationship = sheet.Part.CreateRelationship(UriHelper.ResolvePartUri(sheet.WorksheetUri, PivotTableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");

		_cacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, sourceAddress, tblId);
		_cacheDefinition.Relationship = Part.CreateRelationship(UriHelper.ResolvePartUri(PivotTableUri, _cacheDefinition.CacheDefinitionUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");

		sheet.Workbook.AddPivotTable(CacheID.ToString(), _cacheDefinition.CacheDefinitionUri);

		LoadFields();

		using var r = sheet.Cells[address.Address];
		r.Clear();
	}
	private void init() => SchemaNodeOrder = ["location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "pageItems", "dataFields", "dataItems", "formats", "pivotTableStyleInfo"];
	private void LoadFields()
	{
		var index = 0;
		//Add fields.
		foreach (XmlElement fieldElem in TopNode.SelectNodes("d:pivotFields/d:pivotField", NameSpaceManager))
		{
			var fld = new ExcelPivotTableField(NameSpaceManager, fieldElem, this, index, index++);
			Fields.AddInternal(fld);
		}

		//Add fields.
		index = 0;
		foreach (XmlElement fieldElem in _cacheDefinition.TopNode.SelectNodes("d:cacheFields/d:cacheField", NameSpaceManager))
		{
			var fld = Fields[index++];
			fld.SetCacheFieldNode(fieldElem);
		}
	}
	private static string GetStartXml(string name, int id, ExcelAddressBase address, ExcelAddressBase sourceAddress)
	{
		var xml = new StringBuilder();
		xml.AppendFormat("<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"{0}\" cacheId=\"{1}\" dataOnRows=\"1\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" dataCaption=\"Data\"  createdVersion=\"4\" showMemberPropertyTips=\"0\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" indent=\"0\" compact=\"0\" compactData=\"0\" gridDropZones=\"1\">", ConvertUtil.ExcelEscapeString(name), id);

		xml.AppendFormat("<location ref=\"{0}\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" /> ", address.FirstAddress);
		xml.AppendFormat("<pivotFields count=\"{0}\">", sourceAddress._toCol - sourceAddress._fromCol + 1);
		for (var col = sourceAddress._fromCol; col <= sourceAddress._toCol; col++)
		{
			xml.Append("<pivotField showAll=\"0\" />");
		}

		xml.Append("</pivotFields>");
		xml.Append("<pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\" />");
		xml.Append("</pivotTableDefinition>");
		return xml.ToString();
	}
	internal Packaging.ZipPackagePart Part
	{
		get;
		set;
	}
	/// <summary>
	/// Provides access to the XML data representing the pivottable in the package.
	/// </summary>
	public XmlDocument PivotTableXml { get; private set; }
	/// <summary>
	/// The package internal URI to the pivottable Xml Document.
	/// </summary>
	public Uri PivotTableUri
	{
		get;
		internal set;
	}
	internal Packaging.ZipPackageRelationship Relationship
	{
		get;
		set;
	}
	const string NAME_PATH = "@name";
	const string DISPLAY_NAME_PATH = "@displayName";
	/// <summary>
	/// Name of the pivottable object in Excel
	/// </summary>
	public string Name
	{
		get
		{
			return GetXmlNodeString(NAME_PATH);
		}
		set
		{
			if (WorkSheet.Workbook.ExistsTableName(value))
			{
				throw (new ArgumentException("PivotTable name is not unique"));
			}

			TableNameHelper.RenameTable(WorkSheet, Name, value);
			SetXmlNodeString(NAME_PATH, value);
			SetXmlNodeString(DISPLAY_NAME_PATH, cleanDisplayName(value));
		}
	}
	ExcelPivotCacheDefinition _cacheDefinition;
	/// <summary>
	/// Reference to the pivot table cache definition object
	/// </summary>
	public ExcelPivotCacheDefinition CacheDefinition
	{
		get
		{
			_cacheDefinition ??= new ExcelPivotCacheDefinition(NameSpaceManager, this, null, 1);

			return _cacheDefinition;
		}
	}
	private static string cleanDisplayName(string name) => Regex.Replace(name, @"[^\w\.-_]", "_");
	#region "Internal Properties"
	internal int CacheID
	{
		get
		{
			return GetXmlNodeInt("@cacheId");
		}
		set
		{
			SetXmlNodeString("@cacheId", value.ToString());
		}
	}

	#endregion

}
