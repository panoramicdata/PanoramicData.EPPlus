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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// The collection of worksheets for the workbook
/// </summary>
public partial class ExcelWorksheets : XmlHelper, IEnumerable<ExcelWorksheet>, IDisposable
{
	private ExcelPackage _pck;
	private Dictionary<int, ExcelWorksheet> _worksheets;
	private readonly XmlNamespaceManager _namespaceManager;
	internal ExcelWorksheets(ExcelPackage pck, XmlNamespaceManager nsm, XmlNode topNode) :
		base(nsm, topNode)
	{
		_pck = pck;
		_namespaceManager = nsm;
		_worksheets = [];
		var positionID = _pck._worksheetAdd;

		foreach (XmlNode sheetNode in topNode.ChildNodes)
		{
			if (sheetNode.NodeType == XmlNodeType.Element)
			{
				var name = sheetNode.Attributes["name"].Value;
				//Get the relationship id
				var relId = sheetNode.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships).Value;
				var sheetID = Convert.ToInt32(sheetNode.Attributes["sheetId"].Value);

				//Hidden property
				var hidden = eWorkSheetHidden.Visible;
				XmlNode attr = sheetNode.Attributes["state"];
				if (attr != null)
				{
					hidden = TranslateHidden(attr.Value);
				}

				var sheetRelation = pck.Workbook.Part.GetRelationship(relId);
				var uriWorksheet = UriHelper.ResolvePartUri(pck.Workbook.WorkbookUri, sheetRelation.TargetUri);

				//add the worksheet
				if (sheetRelation.RelationshipType.EndsWith("chartsheet"))
				{
					_worksheets.Add(positionID, new ExcelChartsheet(_namespaceManager, _pck, relId, uriWorksheet, name, sheetID, positionID, hidden));
				}
				else
				{
					_worksheets.Add(positionID, new ExcelWorksheet(_namespaceManager, _pck, relId, uriWorksheet, name, sheetID, positionID, hidden));
				}

				positionID++;
			}
		}
	}

	private static eWorkSheetHidden TranslateHidden(string value) => value switch
	{
		"hidden" => eWorkSheetHidden.Hidden,
		"veryHidden" => eWorkSheetHidden.VeryHidden,
		_ => eWorkSheetHidden.Visible,
	};
	/// <summary>
	/// Returns the number of worksheets in the workbook
	/// </summary>
	public int Count => (_worksheets.Count);
	private const string ERR_DUP_WORKSHEET = "A worksheet with this name already exists in the workbook";
	internal const string WORKSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
	internal const string CHARTSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
	/// <summary>
	/// Foreach support
	/// </summary>
	/// <returns>An enumerator</returns>
	public IEnumerator<ExcelWorksheet> GetEnumerator() => (_worksheets.Values.GetEnumerator());

	IEnumerator IEnumerable.GetEnumerator() => (_worksheets.Values.GetEnumerator());

	/// <summary>
	/// Returns the worksheet at the specified position. 
	/// </summary>
	/// <param name="PositionID">The position of the worksheet. Collection is zero-based or one-base depending on the Package.Compatibility.IsWorksheets1Based propery. Default is Zero based</param>
	/// <seealso cref="ExcelPackage.Compatibility"/>
	/// <returns></returns>
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	public ExcelWorksheet this[int PositionID]
	{
		get
		{
			return _worksheets.TryGetValue(PositionID, out var value)
				? value
				: throw new IndexOutOfRangeException("Worksheet position out of range.");
		}
	}

	/// <summary>
	/// Returns the worksheet matching the specified name
	/// </summary>
	/// <param name="Name">The name of the worksheet</param>
	/// <returns></returns>
	public ExcelWorksheet this[string Name]
	{
		get
		{
			return GetByName(Name);
		}
	}
	internal ExcelWorksheet GetBySheetID(int localSheetID)
	{
		foreach (var ws in this)
		{
			if (ws.SheetID == localSheetID)
			{
				return ws;
			}
		}

		return null;
	}
	private ExcelWorksheet GetByName(string Name)
	{
		if (string.IsNullOrEmpty(Name))
		{
			return null;
		}

		ExcelWorksheet xlWorksheet = null;
		foreach (var worksheet in _worksheets.Values)
		{
			if (worksheet.Name.Equals(Name, StringComparison.OrdinalIgnoreCase))
			{
				xlWorksheet = worksheet;
			}
		}

		return xlWorksheet;
	}

	public void Dispose()
	{
		if (_worksheets != null)
		{
			foreach (var sheet in _worksheets.Values)
			{
				sheet.Dispose();
			}

			_worksheets = null;
			_pck = null;
		}
	}
} // end class Worksheets
