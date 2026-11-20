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
 * Eyal Seagull        Added       		  2012-04-03
 *******************************************************************************/
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.ConditionalFormatting.Rules;

/// <summary>
/// Databar
/// </summary>
public class ExcelConditionalFormattingDataBar
  : ExcelConditionalFormattingRule,
	IExcelConditionalFormattingDataBarGroup
{
	/****************************************************************************************/

	#region Private Properties

	#endregion Private Properties

	/****************************************************************************************/

	#region Constructors
	/// <summary>
	/// 
	/// </summary>
	/// <param name="type"></param>
	/// <param name="priority"></param>
	/// <param name="address"></param>
	/// <param name="worksheet"></param>
	/// <param name="itemElementNode"></param>
	/// <param name="namespaceManager"></param>
	internal ExcelConditionalFormattingDataBar(
	  eExcelConditionalFormattingRuleType type,
	  ExcelAddress address,
	  int priority,
	  ExcelWorksheet worksheet,
	  XmlNode itemElementNode = null,
	  XmlNamespaceManager namespaceManager = null)
		: base(
		  type,
		  address,
		  priority,
		  worksheet,
		  itemElementNode,
		  namespaceManager == null ? worksheet.NameSpaceManager : namespaceManager)
	{
		var s = SchemaNodeOrder;
		Array.Resize(ref s, s.Length + 2);    //Fixes issue 15429. Append node order instead om overwriting it.
		s[^2] = "cfvo";
		s[^1] = "color";
		SchemaNodeOrder = s;

		if (itemElementNode != null && itemElementNode.HasChildNodes)
		{
			LoadValues(type, address, worksheet, itemElementNode, namespaceManager);
		}
		else
		{
			var dataBarNode = CreateComplexNode(
			  Node,
			  ExcelConditionalFormattingConstants.Paths.DataBar);

			LowValue = AppendValue(dataBarNode, eExcelConditionalFormattingValueObjectType.Min, address, priority, worksheet, namespaceManager);
			HighValue = AppendValue(dataBarNode, eExcelConditionalFormattingValueObjectType.Max, address, priority, worksheet, namespaceManager);
		}

		Type = type;
	}

	// The first cfvo is the low value; any later one overwrites the high value.
	private void LoadValues(
	  eExcelConditionalFormattingRuleType type,
	  ExcelAddress address,
	  ExcelWorksheet worksheet,
	  XmlNode itemElementNode,
	  XmlNamespaceManager namespaceManager)
	{
		var high = false;
		foreach (XmlNode node in itemElementNode.SelectNodes("d:dataBar/d:cfvo", NameSpaceManager))
		{
			var value = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, node, namespaceManager);
			if (high)
			{
				HighValue = value;
			}
			else
			{
				LowValue = value;
				high = true;
			}
		}
	}

	private static ExcelConditionalFormattingIconDataBarValue AppendValue(
	  XmlNode dataBarNode,
	  eExcelConditionalFormattingValueObjectType valueType,
	  ExcelAddress address,
	  int priority,
	  ExcelWorksheet worksheet,
	  XmlNamespaceManager namespaceManager)
	{
		var node = dataBarNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
		dataBarNode.AppendChild(node);
		return new ExcelConditionalFormattingIconDataBarValue(
			valueType,
			0,
			"",
			eExcelConditionalFormattingRuleType.DataBar,
			address,
			priority,
			worksheet,
			node,
			namespaceManager);
	}
	#endregion Constructors
	private const string _showValuePath = "d:dataBar/@showValue";
	public bool ShowValue
	{
		get
		{
			return GetXmlNodeBool(_showValuePath, true);
		}
		set
		{
			SetXmlNodeBool(_showValuePath, value);
		}
	}


	public ExcelConditionalFormattingIconDataBarValue LowValue
	{
		get;
		internal set;
	}

	public ExcelConditionalFormattingIconDataBarValue HighValue
	{
		get;
		internal set;
	}


	private const string _colorPath = "d:dataBar/d:color/@rgb";
	public Color Color
	{
		get
		{
			var rgb = GetXmlNodeString(_colorPath);
			return !string.IsNullOrEmpty(rgb) ? Color.FromArgb(int.Parse(rgb, NumberStyles.HexNumber)) : Color.White;
		}
		set
		{
			SetXmlNodeString(_colorPath, value.ToArgb().ToString("X"));
		}
	}
}