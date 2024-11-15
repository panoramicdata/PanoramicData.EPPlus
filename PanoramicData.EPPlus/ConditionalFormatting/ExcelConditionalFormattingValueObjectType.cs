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
 * Author          Change						                  Date
 * ******************************************************************************
 * Eyal Seagull    Conditional Formatting Adaption    2012-04-03
 *******************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

/// <summary>
/// Functions related to the <see cref="ExcelConditionalFormattingColorScaleValue"/>
/// </summary>
internal static class ExcelConditionalFormattingValueObjectType
{
	/// <summary>
	/// Get the sequencial order of a cfvo/color by its position.
	/// </summary>
	/// <param name="position"></param>
	/// <param name="ruleType"></param>
	/// <returns>1, 2 or 3</returns>
	internal static int GetOrderByPosition(
		eExcelConditionalFormattingValueObjectPosition position,
		eExcelConditionalFormattingRuleType ruleType)
	{
		switch (position)
		{
			case eExcelConditionalFormattingValueObjectPosition.Low:
				return 1;

			case eExcelConditionalFormattingValueObjectPosition.Middle:
				return 2;

			case eExcelConditionalFormattingValueObjectPosition.High:
				// Check if the rule type is TwoColorScale.
				if (ruleType == eExcelConditionalFormattingRuleType.TwoColorScale)
				{
					// There are only "Low" and "High". So "High" is the second
					return 2;
				}

				// There are "Low", "Middle" and "High". So "High" is the third
				return 3;
		}

		return 0;
	}

	/// <summary>
	/// Get the CFVO type by its @type attribute
	/// </summary>
	/// <param name="attribute"></param>
	/// <returns></returns>
	public static eExcelConditionalFormattingValueObjectType GetTypeByAttrbiute(
		string attribute) => attribute switch
		{
			ExcelConditionalFormattingConstants.CfvoType.Min => eExcelConditionalFormattingValueObjectType.Min,
			ExcelConditionalFormattingConstants.CfvoType.Max => eExcelConditionalFormattingValueObjectType.Max,
			ExcelConditionalFormattingConstants.CfvoType.Num => eExcelConditionalFormattingValueObjectType.Num,
			ExcelConditionalFormattingConstants.CfvoType.Formula => eExcelConditionalFormattingValueObjectType.Formula,
			ExcelConditionalFormattingConstants.CfvoType.Percent => eExcelConditionalFormattingValueObjectType.Percent,
			ExcelConditionalFormattingConstants.CfvoType.Percentile => eExcelConditionalFormattingValueObjectType.Percentile,
			_ => throw new Exception(
				ExcelConditionalFormattingConstants.Errors.UnexistentCfvoTypeAttribute),
		};

	/// <summary>
	/// 
	/// </summary>
	/// <param name="position"></param>
	///<param name="ruleType"></param>
	/// <param name="topNode"></param>
	/// <param name="nameSpaceManager"></param>
	/// <returns></returns>
	public static XmlNode GetCfvoNodeByPosition(
		eExcelConditionalFormattingValueObjectPosition position,
		eExcelConditionalFormattingRuleType ruleType,
		XmlNode topNode,
		XmlNamespaceManager nameSpaceManager)
	{
		// Get the corresponding <cfvo> node (by the position)
		var node = topNode.SelectSingleNode(
			string.Format(
				"{0}[position()={1}]",
				// {0}
				ExcelConditionalFormattingConstants.Paths.Cfvo,
				// {1}
				GetOrderByPosition(position, ruleType)),
			nameSpaceManager);

		return node == null
			?         throw new Exception(
	  ExcelConditionalFormattingConstants.Errors.MissingCfvoNode)
			: node;
	}

	/// <summary>
	/// 
	/// </summary>
	/// <param name="type"></param>
	/// <returns></returns>
	public static string GetAttributeByType(
		eExcelConditionalFormattingValueObjectType type) => type switch
		{
			eExcelConditionalFormattingValueObjectType.Min => ExcelConditionalFormattingConstants.CfvoType.Min,
			eExcelConditionalFormattingValueObjectType.Max => ExcelConditionalFormattingConstants.CfvoType.Max,
			eExcelConditionalFormattingValueObjectType.Num => ExcelConditionalFormattingConstants.CfvoType.Num,
			eExcelConditionalFormattingValueObjectType.Formula => ExcelConditionalFormattingConstants.CfvoType.Formula,
			eExcelConditionalFormattingValueObjectType.Percent => ExcelConditionalFormattingConstants.CfvoType.Percent,
			eExcelConditionalFormattingValueObjectType.Percentile => ExcelConditionalFormattingConstants.CfvoType.Percentile,
			_ => string.Empty,
		};

	/// <summary>
	/// Get the cfvo (§18.3.1.11) node parent by the rule type. Can be any of the following:
	/// "colorScale" (§18.3.1.16); "dataBar" (§18.3.1.28); "iconSet" (§18.3.1.49)
	/// </summary>
	/// <param name="ruleType"></param>
	/// <returns></returns>
	public static string GetParentPathByRuleType(
		eExcelConditionalFormattingRuleType ruleType) => ruleType switch
		{
			eExcelConditionalFormattingRuleType.TwoColorScale or eExcelConditionalFormattingRuleType.ThreeColorScale => ExcelConditionalFormattingConstants.Paths.ColorScale,
			eExcelConditionalFormattingRuleType.ThreeIconSet or eExcelConditionalFormattingRuleType.FourIconSet or eExcelConditionalFormattingRuleType.FiveIconSet => ExcelConditionalFormattingConstants.Paths.IconSet,
			eExcelConditionalFormattingRuleType.DataBar => ExcelConditionalFormattingConstants.Paths.DataBar,
			_ => string.Empty,
		};

	/// <summary>
	/// 
	/// </summary>
	/// <param name="nodeType"></param>
	/// <returns></returns>
	public static string GetNodePathByNodeType(
		eExcelConditionalFormattingValueObjectNodeType nodeType) => nodeType switch
		{
			eExcelConditionalFormattingValueObjectNodeType.Cfvo => ExcelConditionalFormattingConstants.Paths.Cfvo,
			eExcelConditionalFormattingValueObjectNodeType.Color => ExcelConditionalFormattingConstants.Paths.Color,
			_ => string.Empty,
		};
}