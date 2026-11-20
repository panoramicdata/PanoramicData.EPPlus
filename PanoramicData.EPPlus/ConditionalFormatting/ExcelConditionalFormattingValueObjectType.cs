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
		eExcelConditionalFormattingRuleType ruleType) => position switch
		{
			eExcelConditionalFormattingValueObjectPosition.Low => 1,
			eExcelConditionalFormattingValueObjectPosition.Middle => 2,
			// A two colour scale has no middle value, so High is the second node
			eExcelConditionalFormattingValueObjectPosition.High => ruleType == eExcelConditionalFormattingRuleType.TwoColorScale ? 2 : 3,
			_ => 0,
		};

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
			_ => throw new InvalidOperationException(
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
		var node = topNode.SelectSingleNode(
			string.Format("{0}[position()={1}]", ExcelConditionalFormattingConstants.Paths.Cfvo, GetOrderByPosition(position, ruleType)),
			nameSpaceManager);

		return node ?? throw new InvalidOperationException(ExcelConditionalFormattingConstants.Errors.MissingCfvoNode);
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

	/// <summary>
	/// Finds the existing cfvo parent node of the rule with this address and priority.
	/// </summary>
	internal static XmlNode GetCfvoParentNode(
		ExcelWorksheet worksheet,
		eExcelConditionalFormattingRuleType ruleType,
		ExcelAddress address,
		int priority)
	{
		var parentNodePath = GetParentPathByRuleType(ruleType);
		if (parentNodePath == string.Empty)
		{
			throw new InvalidOperationException(
				ExcelConditionalFormattingConstants.Errors.MissingCfvoParentNode);
		}

		var node = worksheet.WorksheetXml.SelectSingleNode(
			string.Format(
				"//{0}[{1}='{2}']/{3}[{4}='{5}']/{6}",
				ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
				ExcelConditionalFormattingConstants.Paths.SqrefAttribute,
				address.Address,
				ExcelConditionalFormattingConstants.Paths.CfRule,
				ExcelConditionalFormattingConstants.Paths.PriorityAttribute,
				priority,
				parentNodePath),
			worksheet.NameSpaceManager);

		return node ?? throw new InvalidOperationException(
			ExcelConditionalFormattingConstants.Errors.MissingCfvoParentNode);
	}
}