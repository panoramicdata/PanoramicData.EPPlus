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
using OfficeOpenXml.Utils;
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

/// <summary>
/// 18.3.1.11 cfvo (Conditional Format Value Object)
/// Describes the values of the interpolation points in a gradient scale.
/// </summary>
public class ExcelConditionalFormattingColorScaleValue
	: XmlHelper
{
	/****************************************************************************************/

	#region Private Properties
	private eExcelConditionalFormattingValueObjectPosition _position;
	private eExcelConditionalFormattingRuleType _ruleType;
	private readonly ExcelWorksheet _worksheet;
	#endregion Private Properties

	/****************************************************************************************/

	#region Constructors
	/// <summary>
	/// Initialize the cfvo (§18.3.1.11) node
	/// </summary>
	/// <param name="position"></param>
	/// <param name="type"></param>
	/// <param name="color"></param>
	/// <param name="value"></param>
	/// <param name="formula"></param>
	/// <param name="ruleType"></param>
	/// <param name="address"></param>
	/// <param name="priority"></param>
	/// <param name="worksheet"></param>
	/// <param name="itemElementNode">The cfvo (§18.3.1.11) node parent. Can be any of the following:
	/// colorScale (§18.3.1.16); dataBar (§18.3.1.28); iconSet (§18.3.1.49)</param>
	/// <param name="namespaceManager"></param>
	internal ExcelConditionalFormattingColorScaleValue(
		eExcelConditionalFormattingValueObjectPosition position,
		eExcelConditionalFormattingValueObjectType type,
		Color color,
		double value,
		string formula,
		eExcelConditionalFormattingRuleType ruleType,
  ExcelAddress address,
  int priority,
		ExcelWorksheet worksheet,
		XmlNode itemElementNode,
		XmlNamespaceManager namespaceManager)
		: base(
			namespaceManager,
			itemElementNode)
	{
		Require.Argument(priority).IsInRange(1, int.MaxValue, "priority");
		Require.Argument(address).IsNotNull("address");
		Require.Argument(worksheet).IsNotNull("worksheet");

		// Save the worksheet for private methods to use
		_worksheet = worksheet;

		// Schema order list
		SchemaNodeOrder =
		[
	ExcelConditionalFormattingConstants.Nodes.Cfvo,
	ExcelConditionalFormattingConstants.Nodes.Color
		];

		itemElementNode ??= ExcelConditionalFormattingValueObjectType.GetCfvoParentNode(_worksheet, ruleType, address, priority);

		// Point to the <cfvo> parent node (<colorScale>, <dataBar> or <iconSet>)
		// This is different than normal, as TopNode does not point to the node itself but to
		// its PARENT. Later, in the CreateNodeByOrdem method the TopNode will be updated.
		TopNode = itemElementNode;

		// Save the attributes
		Position = position;
		RuleType = ruleType;
		Type = type;
		Color = color;
		Value = value;
		Formula = formula;
	}

	/// <summary>
	/// Initialize the <see cref="ExcelConditionalFormattingColorScaleValue"/>
	/// </summary>
	/// <param name="position"></param>
	/// <param name="type"></param>
	/// <param name="color"></param>
	/// <param name="value"></param>
	/// <param name="formula"></param>
	/// <param name="ruleType"></param>
	/// <param name="priority"></param>
	/// <param name="address"></param>
	/// <param name="worksheet"></param>
	/// <param name="namespaceManager"></param>
	internal ExcelConditionalFormattingColorScaleValue(
		eExcelConditionalFormattingValueObjectPosition position,
		eExcelConditionalFormattingValueObjectType type,
		Color color,
		double value,
		string formula,
		eExcelConditionalFormattingRuleType ruleType,
  ExcelAddress address,
  int priority,
		ExcelWorksheet worksheet,
		XmlNamespaceManager namespaceManager)
		: this(
			position,
			type,
			color,
			value,
			formula,
			ruleType,
	address,
	priority,
			worksheet,
			null,
			namespaceManager)
	{
	}

	/// <summary>
	/// Initialize the <see cref="ExcelConditionalFormattingColorScaleValue"/>
	/// </summary>
	/// <param name="position"></param>
	/// <param name="type"></param>
	/// <param name="color"></param>
	/// <param name="ruleType"></param>
	/// <param name="priority"></param>
	/// <param name="address"></param>
	/// <param name="worksheet"></param>
	/// <param name="namespaceManager"></param>
	internal ExcelConditionalFormattingColorScaleValue(
		eExcelConditionalFormattingValueObjectPosition position,
		eExcelConditionalFormattingValueObjectType type,
		Color color,
		eExcelConditionalFormattingRuleType ruleType,
  ExcelAddress address,
  int priority,
		ExcelWorksheet worksheet,
		XmlNamespaceManager namespaceManager)
		: this(
			position,
			type,
			color,
			0,
			null,
			ruleType,
	address,
	priority,
			worksheet,
			null,
			namespaceManager)
	{
	}
	internal static ExcelConditionalFormattingColorScaleValue CreateDefaultLow(
		eExcelConditionalFormattingRuleType ruleType,
		ExcelAddress address,
		int priority,
		ExcelWorksheet worksheet,
		XmlNamespaceManager namespaceManager) => new(
			eExcelConditionalFormattingValueObjectPosition.Low,
			eExcelConditionalFormattingValueObjectType.Min,
			ExcelConditionalFormattingConstants.Colors.CfvoLowValue,
			ruleType,
			address,
			priority,
			worksheet,
			namespaceManager);

	internal static ExcelConditionalFormattingColorScaleValue CreateDefaultHigh(
		eExcelConditionalFormattingRuleType ruleType,
		ExcelAddress address,
		int priority,
		ExcelWorksheet worksheet,
		XmlNamespaceManager namespaceManager) => new(
			eExcelConditionalFormattingValueObjectPosition.High,
			eExcelConditionalFormattingValueObjectType.Max,
			ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
			ruleType,
			address,
			priority,
			worksheet,
			namespaceManager);
	#endregion Constructors

	/****************************************************************************************/

	#region Methods
	/// <summary>
	/// Get the node order (1, 2 ou 3) according to the Position (Low, Middle and High)
	/// and the Rule Type (TwoColorScale ou ThreeColorScale).
	/// </summary>
	/// <returns></returns>
	private int GetNodeOrder() => ExcelConditionalFormattingValueObjectType.GetOrderByPosition(
			Position,
			RuleType);

	private static string GetOrderedNodePath(string nodePath, int nodeOrder) => string.Format("{0}[position()={1}]", nodePath, nodeOrder);

	private string GetOrderedAttributePath(string nodePath, string attributePath) => GetOrderedNodePath(nodePath, GetNodeOrder()) + "/" + attributePath;

	/// <summary>
	/// Create the 'cfvo'/'color' nodes in the right order. They should appear like this:
	///		"cfvo"   --> Low Value (value object)
	///		"cfvo"   --> Middle Value (value object)
	///		"cfvo"   --> High Value (value object)
	///		"color"  --> Low Value (color)
	///		"color"  --> Middle Value (color)
	///		"color"  --> High Value (color)
	/// </summary>
	/// <param name="nodeType"></param>
	/// <param name="attributePath"></param>
	/// <param name="attributeValue"></param>
	private void CreateNodeByOrdem(
		eExcelConditionalFormattingValueObjectNodeType nodeType,
		string attributePath,
		string attributeValue)
	{
		var currentTopNode = TopNode;

		var nodePath = ExcelConditionalFormattingValueObjectType.GetNodePathByNodeType(nodeType);
		var nodeOrder = GetNodeOrder();
		var nodeInsertOrder = eNodeInsertOrder.SchemaOrder;
		XmlNode referenceNode = null;

		if (nodeOrder > 1)
		{
			referenceNode = TopNode.SelectSingleNode(
				GetOrderedNodePath(nodePath, nodeOrder - 1),
				_worksheet.NameSpaceManager);

			if (referenceNode != null)
			{
				nodeInsertOrder = eNodeInsertOrder.After;
			}
		}

		var node = CreateComplexNode(
			TopNode,
			GetOrderedNodePath(nodePath, nodeOrder),
			nodeInsertOrder,
			referenceNode);

		// The XmlHelper functions work relative to TopNode, so point it at the new node while setting the attribute
		TopNode = node;

		// An empty attributeValue removes the attribute
		SetXmlNodeString(
		  node,
		  attributePath,
		  attributeValue,
		  true);

		TopNode = currentTopNode;
	}
	#endregion Methos

	/****************************************************************************************/

	#region Exposed Properties
	/// <summary>
	///
	/// </summary>
	internal eExcelConditionalFormattingValueObjectPosition Position
	{
		get { return _position; }
		set { _position = value; }
	}

	/// <summary>
	///
	/// </summary>
	internal eExcelConditionalFormattingRuleType RuleType
	{
		get { return _ruleType; }
		set { _ruleType = value; }
	}

	/// <summary>
	///
	/// </summary>
	public eExcelConditionalFormattingValueObjectType Type
	{
		get
		{
			var typeAttribute = GetXmlNodeString(
				GetOrderedAttributePath(
					ExcelConditionalFormattingConstants.Paths.Cfvo,
					ExcelConditionalFormattingConstants.Paths.TypeAttribute));

			return ExcelConditionalFormattingValueObjectType.GetTypeByAttrbiute(typeAttribute);
		}
		set
		{
			CreateNodeByOrdem(
				eExcelConditionalFormattingValueObjectNodeType.Cfvo,
				ExcelConditionalFormattingConstants.Paths.TypeAttribute,
				ExcelConditionalFormattingValueObjectType.GetAttributeByType(value));

			// Min and Max have no @val, so remove any left over from the previous type
			if (Type is eExcelConditionalFormattingValueObjectType.Min or eExcelConditionalFormattingValueObjectType.Max)
			{
				var nodePath = ExcelConditionalFormattingValueObjectType.GetNodePathByNodeType(
					  eExcelConditionalFormattingValueObjectNodeType.Cfvo);

				// Setting the attribute to '' removes it
				CreateComplexNode(
				  TopNode,
				  GetOrderedAttributePath(nodePath, ExcelConditionalFormattingConstants.Paths.ValAttribute) + "=''");
			}
		}
	}

	/// <summary>
	///
	/// </summary>
	public Color Color
	{
		get
		{
			// Color Code like "FF5B34F2"
			var colorCode = GetXmlNodeString(
				GetOrderedAttributePath(
					ExcelConditionalFormattingConstants.Paths.Color,
					ExcelConditionalFormattingConstants.Paths.RgbAttribute));

			return ExcelConditionalFormattingHelper.ConvertFromColorCode(colorCode);
		}
		set
		{
			// Use the color code to store (Ex. "FF5B35F2")
			CreateNodeByOrdem(
				eExcelConditionalFormattingValueObjectNodeType.Color,
				ExcelConditionalFormattingConstants.Paths.RgbAttribute,
				value.ToArgb().ToString("x"));
		}
	}

	/// <summary>
	/// Get/Set the 'cfvo' node @val attribute
	/// </summary>
	public Double Value
	{
		get
		{
			return GetXmlNodeDouble(
				GetOrderedAttributePath(
					ExcelConditionalFormattingConstants.Paths.Cfvo,
					ExcelConditionalFormattingConstants.Paths.ValAttribute));
		}
		set
		{
			var valueToStore = string.Empty;

			// Only some types use the @val attribute
			if (Type is eExcelConditionalFormattingValueObjectType.Num or eExcelConditionalFormattingValueObjectType.Percent or eExcelConditionalFormattingValueObjectType.Percentile)
			{
				valueToStore = value.ToString(CultureInfo.InvariantCulture);
			}

			CreateNodeByOrdem(
				eExcelConditionalFormattingValueObjectNodeType.Cfvo,
				ExcelConditionalFormattingConstants.Paths.ValAttribute,
				valueToStore);
		}
	}

	/// <summary>
	/// Get/Set the Formula of the Object Value (uses the same attribute as the Value)
	/// </summary>
	public string Formula
	{
		get
		{
			if (Type != eExcelConditionalFormattingValueObjectType.Formula)
			{
				return string.Empty;
			}

			// Excel stores the formula in the @val attribute
			return GetXmlNodeString(
				GetOrderedAttributePath(
					ExcelConditionalFormattingConstants.Paths.Cfvo,
					ExcelConditionalFormattingConstants.Paths.ValAttribute));
		}
		set
		{
			if (Type == eExcelConditionalFormattingValueObjectType.Formula)
			{
				CreateNodeByOrdem(
					eExcelConditionalFormattingValueObjectNodeType.Cfvo,
					ExcelConditionalFormattingConstants.Paths.ValAttribute,
					value ?? string.Empty);
			}
		}
	}
	#endregion Exposed Properties

	/****************************************************************************************/
}