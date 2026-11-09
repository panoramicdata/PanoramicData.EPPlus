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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;


/// <summary>
/// A pivot table field.
/// </summary>
public partial class ExcelPivotTableField : XmlHelper
{
	internal ExcelPivotTable _table;
	internal ExcelPivotTableField(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTable table, int index, int baseIndex) :
		base(ns, topNode)
	{
		Index = index;
		BaseIndex = baseIndex;
		_table = table;
	}
	public int Index
	{
		get;
		set;
	}
	internal int BaseIndex
	{
		get;
		set;
	}
	/// <summary>
	/// Name of the field
	/// </summary>
	public string Name
	{
		get
		{
			var v = GetXmlNodeString("@name");
			if (v == "")
			{
				return _cacheFieldHelper.GetXmlNodeString("@name");
			}

			return v;
		}
		set
		{
			SetXmlNodeString("@name", value);
		}
	}
	/// <summary>
	/// Compact mode
	/// </summary>
	public bool Compact
	{
		get
		{
			return GetXmlNodeBool("@compact");
		}
		set
		{
			SetXmlNodeBool("@compact", value);
		}
	}
	/// <summary>
	/// A boolean that indicates whether the items in this field should be shown in Outline form
	/// </summary>
	public bool Outline
	{
		get
		{
			return GetXmlNodeBool("@outline");
		}
		set
		{
			SetXmlNodeBool("@outline", value);
		}
	}
	/// <summary>
	/// The custom text that is displayed for the subtotals label
	/// </summary>
	public bool SubtotalTop
	{
		get
		{
			return GetXmlNodeBool("@subtotalTop");
		}
		set
		{
			SetXmlNodeBool("@subtotalTop", value);
		}
	}
	/// <summary>
	/// Indicates whether the field can have multiple items selected in the page field
	/// </summary>
	public bool MultipleItemSelectionAllowed
	{
		get
		{
			return GetXmlNodeBool("@multipleItemSelectionAllowed");
		}
		set
		{
			SetXmlNodeBool("@multipleItemSelectionAllowed", value);
		}
	}
	/// <summary>
	/// The type of sort that is applied to this field
	/// </summary>
	public eSortType Sort
	{
		get
		{
			var v = GetXmlNodeString("@sortType");
			return v == "" ? eSortType.None : Enum.Parse<eSortType>(v, true);
		}
		set
		{
			if (value == eSortType.None)
			{
				DeleteNode("@sortType");
			}
			else
			{
				SetXmlNodeString("@sortType", value.ToString().ToLower(CultureInfo.InvariantCulture));
			}
		}
	}
	/// <summary>
	/// A boolean that indicates whether manual filter is in inclusive mode
	/// </summary>
	public bool IncludeNewItemsInFilter
	{
		get
		{
			return GetXmlNodeBool("@includeNewItemsInFilter");
		}
		set
		{
			SetXmlNodeBool("@includeNewItemsInFilter", value);
		}
	}
	/// <summary>
	/// Enumeration of the different subtotal operations that can be applied to page, row or column fields
	/// </summary>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Major Code Smell", "S2372:Exceptions should not be thrown from property getters", Justification = "Public API: an unparsable subtotal in the XML is reported when the property is read.")]
	public eSubTotalFunctions SubTotalFunctions
	{
		get
		{
			eSubTotalFunctions ret = 0;
			var nl = TopNode.SelectNodes("d:items/d:item/@t", NameSpaceManager);
			if (nl.Count == 0)
			{
				return eSubTotalFunctions.None;
			}

			foreach (XmlAttribute item in nl)
			{
				try
				{
					ret |= Enum.Parse<eSubTotalFunctions>(item.Value, true);
				}
				catch (ArgumentException ex)
				{
					throw new ArgumentException("Unable to parse value of " + item.Value + " to a valid pivot table subtotal function", ex);
				}
			}

			return ret;
		}
		set
		{
			if ((value & eSubTotalFunctions.None) == eSubTotalFunctions.None && (value != eSubTotalFunctions.None))
			{
				throw (new ArgumentException("Value None can not be combined with other values."));
			}

			if ((value & eSubTotalFunctions.Default) == eSubTotalFunctions.Default && (value != eSubTotalFunctions.Default))
			{
				throw (new ArgumentException("Value Default can not be combined with other values."));
			}


			// remove old attribute                 
			var nl = TopNode.SelectNodes("d:items/d:item/@t", NameSpaceManager);
			if (nl.Count > 0)
			{
				foreach (XmlAttribute item in nl)
				{
					DeleteNode("@" + item.Value + "Subtotal");
					item.OwnerElement.ParentNode.RemoveChild(item.OwnerElement);
				}
			}


			if (value == eSubTotalFunctions.None)
			{
				// for no subtotals, set defaultSubtotal to off
				SetXmlNodeBool("@defaultSubtotal", false);
				TopNode.InnerXml = "";
			}
			else
			{
				var innerXml = new StringBuilder();
				var count = 0;
				foreach (var e in Enum.GetValues<eSubTotalFunctions>())
				{
					if ((value & e) == e)
					{
						var newTotalType = e.ToString();
						var totalType = char.ToLowerInvariant(newTotalType[0]) + newTotalType[1..];
						// add new attribute
						SetXmlNodeBool("@" + totalType + "Subtotal", true);
						innerXml.Append("<item t=\"").Append(totalType).Append("\" />");
						count++;
					}
				}

				TopNode.InnerXml = string.Format("<items count=\"{0}\">{1}</items>", count, innerXml);
			}
		}
	}
	/// <summary>
	/// Type of axis
	/// </summary>
	public ePivotFieldAxis Axis
	{
		get
		{
			return GetXmlNodeString("@axis") switch
			{
				"axisRow" => ePivotFieldAxis.Row,
				"axisCol" => ePivotFieldAxis.Column,
				"axisPage" => ePivotFieldAxis.Page,
				"axisValues" => ePivotFieldAxis.Values,
				_ => ePivotFieldAxis.None,
			};
		}
		internal set
		{
			switch (value)
			{
				case ePivotFieldAxis.Row:
					SetXmlNodeString("@axis", "axisRow");
					break;
				case ePivotFieldAxis.Column:
					SetXmlNodeString("@axis", "axisCol");
					break;
				case ePivotFieldAxis.Values:
					SetXmlNodeString("@axis", "axisValues");
					break;
				case ePivotFieldAxis.Page:
					SetXmlNodeString("@axis", "axisPage");
					break;
				default:
					DeleteNode("@axis");
					break;
			}
		}
	}
	/// <summary>
	/// If the field is a row field
	/// </summary>
	public bool IsRowField
	{
		get
		{
			return (TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", Index), NameSpaceManager) != null);
		}
		internal set
		{
			if (value)
			{
				var rowsNode = TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);
				if (rowsNode == null)
				{
					_table.CreateNode("d:rowFields");
				}

				rowsNode = TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);

				AppendField(rowsNode, Index, "field", "x");
				if (BaseIndex == Index)
				{
					TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
				}
				else
				{
					TopNode.InnerXml = "<items count=\"0\"></items>";
				}
			}
			else
			{
				var node = TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", Index), NameSpaceManager) as XmlElement;
				node?.ParentNode.RemoveChild(node);
			}
		}
	}
	/// <summary>
	/// If the field is a column field
	/// </summary>
	public bool IsColumnField
	{
		get
		{
			return (TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", Index), NameSpaceManager) != null);
		}
		internal set
		{
			if (value)
			{
				var columnsNode = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);
				if (columnsNode == null)
				{
					_table.CreateNode("d:colFields");
				}

				columnsNode = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);

				AppendField(columnsNode, Index, "field", "x");
				if (BaseIndex == Index)
				{
					TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
				}
				else
				{
					TopNode.InnerXml = "<items count=\"0\"></items>";
				}
			}
			else
			{
				var node = TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", Index), NameSpaceManager) as XmlElement;
				node?.ParentNode.RemoveChild(node);
			}
		}
	}
	/// <summary>
	/// If the field is a datafield
	/// </summary>
	public bool IsDataField => GetXmlNodeBool("@dataField", false);
	/// <summary>
	/// If the field is a page field.
	/// </summary>
	public bool IsPageField
	{
		get
		{
			return (Axis == ePivotFieldAxis.Page);
		}
		internal set
		{
			if (value)
			{
				var dataFieldsNode = TopNode.SelectSingleNode("../../d:pageFields", NameSpaceManager);
				if (dataFieldsNode == null)
				{
					_table.CreateNode("d:pageFields");
					dataFieldsNode = TopNode.SelectSingleNode("../../d:pageFields", NameSpaceManager);
				}

				TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";

				var node = AppendField(dataFieldsNode, Index, "pageField", "fld");
				_pageFieldSettings = new ExcelPivotTablePageFieldSettings(NameSpaceManager, node, this, Index);
			}
			else
			{
				_pageFieldSettings = null;
				var node = TopNode.SelectSingleNode(string.Format("../../d:pageFields/d:pageField[@fld={0}]", Index), NameSpaceManager) as XmlElement;
				node?.ParentNode.RemoveChild(node);
			}
		}
	}
	internal ExcelPivotTablePageFieldSettings _pageFieldSettings;
	public ExcelPivotTablePageFieldSettings PageFieldSettings => _pageFieldSettings;
	internal eDateGroupBy DateGrouping
	{
		get;
		set;
	}
	ExcelPivotTableFieldGroup _grouping;
	/// <summary>
	/// Grouping settings. 
	/// Null if the field has no grouping otherwise ExcelPivotTableFieldNumericGroup or ExcelPivotTableFieldNumericGroup.
	/// </summary>        
	public ExcelPivotTableFieldGroup Grouping => _grouping;
	#region Private & internal Methods
	internal static XmlElement AppendField(XmlNode rowsNode, int index, string fieldNodeText, string indexAttrText)
	{
		XmlElement prevField = null, newElement;
		foreach (XmlElement field in rowsNode.ChildNodes)
		{
			var x = field.GetAttribute(indexAttrText);
			if (int.TryParse(x, out var fieldIndex))
			{
				if (fieldIndex == index)    //Row already exists
				{
					return field;
				}
			}

			prevField = field;
		}

		newElement = rowsNode.OwnerDocument.CreateElement(fieldNodeText, ExcelPackage.schemaMain);
		newElement.SetAttribute(indexAttrText, index.ToString());
		rowsNode.InsertAfter(newElement, prevField);

		return newElement;
	}
	internal XmlHelperInstance _cacheFieldHelper;
	internal void SetCacheFieldNode(XmlNode cacheField)
	{
		_cacheFieldHelper = new XmlHelperInstance(NameSpaceManager, cacheField);
		var groupNode = cacheField.SelectSingleNode("d:fieldGroup", NameSpaceManager);
		if (groupNode != null)
		{
			var groupBy = groupNode.SelectSingleNode("d:rangePr/@groupBy", NameSpaceManager);
			if (groupBy == null)
			{
				_grouping = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, cacheField);
			}
			else
			{
				DateGrouping = Enum.Parse<eDateGroupBy>(groupBy.Value, true);
				_grouping = new ExcelPivotTableFieldDateGroup(NameSpaceManager, groupNode);
			}
		}
	}
	#endregion
	internal ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem> _items;
	/// <summary>
	/// Pivottable field Items. Used for grouping.
	/// </summary>
	public ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem> Items
	{
		get
		{
			if (_items == null)
			{
				_items = new ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem>(_table);
				foreach (XmlNode node in TopNode.SelectNodes("d:items//d:item", NameSpaceManager))
				{
					var item = new ExcelPivotTableFieldItem(NameSpaceManager, node, this);
					if (item.T == "")
					{
						_items.AddInternal(item);
					}
				}
			}

			return _items;
		}
	}
	/// <summary>
	/// Add numberic grouping to the field
	/// </summary>
	/// <param name="Start">Start value</param>
	/// <param name="End">End value</param>
	/// <param name="Interval">Interval</param>
	public void AddNumericGrouping(double Start, double End, double Interval)
	{
		ValidateGrouping();
		SetNumericGroup(Start, End, Interval);
	}
	/// <summary>
	/// Add a date grouping on this field.
	/// </summary>
	/// <param name="groupBy">Group by</param>
	public void AddDateGrouping(eDateGroupBy groupBy) => AddDateGrouping(groupBy, DateTime.MinValue, DateTime.MaxValue, 1);
	/// <summary>
	/// Add a date grouping on this field.
	/// </summary>
	/// <param name="groupBy">Group by</param>
	/// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
	/// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
	public void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate) => AddDateGrouping(groupBy, startDate, endDate, 1);
	/// <summary>
	/// Add a date grouping on this field.
	/// </summary>
	/// <param name="days">Number of days when grouping on days</param>
	/// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
	/// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
	public void AddDateGrouping(int days, DateTime startDate, DateTime endDate) => AddDateGrouping(eDateGroupBy.Days, startDate, endDate, days);
	static readonly eDateGroupBy[] _dateGroupOrder = [eDateGroupBy.Seconds, eDateGroupBy.Minutes, eDateGroupBy.Hours, eDateGroupBy.Days, eDateGroupBy.Months, eDateGroupBy.Quarters, eDateGroupBy.Years];

	private static void ValidateGroupInterval(eDateGroupBy groupBy, int groupInterval)
	{
		if (groupInterval is < 1 or >= short.MaxValue)
		{
			throw (new ArgumentOutOfRangeException("Group interval is out of range"));
		}

		if (groupInterval > 1 && groupBy != eDateGroupBy.Days)
		{
			throw (new ArgumentException("Group interval is can only be used when groupBy is Days"));
		}
	}

	private void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int groupInterval)
	{
		ValidateGroupInterval(groupBy, groupInterval);
		ValidateGrouping();

		var firstField = true;
		List<ExcelPivotTableField> fields = [];
		foreach (var part in _dateGroupOrder)
		{
			if ((groupBy & part) == part)
			{
				fields.Add(AddField(part, startDate, endDate, ref firstField, part == eDateGroupBy.Days ? groupInterval : 1));
			}
		}

		if (fields.Count > 1)
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/@par", (_table.Fields.Count - 1).ToString());
		}

		if (groupInterval != 1)
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@groupInterval", groupInterval.ToString());
		}
		else
		{
			_cacheFieldHelper.DeleteNode("d:fieldGroup/d:rangePr/@groupInterval");
		}

		_items = null;
	}

	private void ValidateGrouping()
	{
		if (!(IsColumnField || IsRowField))
		{
			throw (new Exception("Field must be a row or column field"));
		}

		foreach (var field in _table.Fields)
		{
			if (field.Grouping != null)
			{
				throw (new Exception("Grouping already exists"));
			}
		}
	}
	private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField, int interval)
	{
		if (firstField == false)
		{
			return AddDateGroupField(groupBy, startDate, endDate, interval);
		}
		else
		{
			firstField = false;
			DateGrouping = groupBy;
			Compact = false;
			SetDateGroup(groupBy, startDate, endDate, interval);
			return this;
		}
	}

	private ExcelPivotTableField AddDateGroupField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int interval)
	{
		//Pivot field
		var topNode = _table.PivotTableXml.SelectSingleNode("//d:pivotFields", _table.NameSpaceManager);
		var fieldNode = _table.PivotTableXml.CreateElement("pivotField", ExcelPackage.schemaMain);
		fieldNode.SetAttribute("compact", "0");
		fieldNode.SetAttribute("outline", "0");
		fieldNode.SetAttribute("showAll", "0");
		fieldNode.SetAttribute("defaultSubtotal", "0");
		topNode.AppendChild(fieldNode);

		var field = new ExcelPivotTableField(_table.NameSpaceManager, fieldNode, _table, _table.Fields.Count, Index)
		{
			DateGrouping = groupBy
		};

		var index = InsertFieldReference(field);

		if (IsRowField)
		{
			_table.RowFields.Insert(field, index);
		}
		else
		{
			_table.ColumnFields.Insert(field, index);
		}

		_table.Fields.AddInternal(field);

		AddCacheField(field, startDate, endDate, interval);
		return field;
	}

	// Inserts the new field before this field in the row or column fields and returns its position.
	private int InsertFieldReference(ExcelPivotTableField field)
	{
		XmlNode rowColFields;
		if (IsRowField)
		{
			rowColFields = TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);
		}
		else
		{
			rowColFields = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);
		}

		var index = 0;
		foreach (XmlElement rowfield in rowColFields.ChildNodes)
		{
			if (int.TryParse(rowfield.GetAttribute("x"), out var fieldIndex) && _table.Fields[fieldIndex].BaseIndex == BaseIndex)
			{
				var newElement = rowColFields.OwnerDocument.CreateElement("field", ExcelPackage.schemaMain);
				newElement.SetAttribute("x", field.Index.ToString());
				rowColFields.InsertBefore(newElement, rowfield);
				break;
			}

			index++;
		}

		return index;
	}
	private void AddCacheField(ExcelPivotTableField field, DateTime startDate, DateTime endDate, int interval)
	{
		//Add Cache definition field.
		var cacheTopNode = _table.CacheDefinition.CacheDefinitionXml.SelectSingleNode("//d:cacheFields", _table.NameSpaceManager);
		var cacheFieldNode = _table.CacheDefinition.CacheDefinitionXml.CreateElement("cacheField", ExcelPackage.schemaMain);

		cacheFieldNode.SetAttribute("name", field.DateGrouping.ToString());
		cacheFieldNode.SetAttribute("databaseField", "0");
		cacheTopNode.AppendChild(cacheFieldNode);
		field.SetCacheFieldNode(cacheFieldNode);

		field.SetDateGroup(field.DateGrouping, startDate, endDate, interval);
	}
}
