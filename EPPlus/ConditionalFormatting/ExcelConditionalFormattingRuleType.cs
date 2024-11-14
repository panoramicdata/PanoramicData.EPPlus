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
/// Functions related to the ExcelConditionalFormattingRule
/// </summary>
internal static class ExcelConditionalFormattingRuleType
{
	/// <summary>
	/// 
	/// </summary>
	/// <param name="attribute"></param>
	/// <param name="topNode"></param>
	/// <param name="nameSpaceManager"></param>
	/// <returns></returns>
	internal static eExcelConditionalFormattingRuleType GetTypeByAttrbiute(
	  string attribute,
	  XmlNode topNode,
	  XmlNamespaceManager nameSpaceManager) => attribute switch
	  {
		  ExcelConditionalFormattingConstants.RuleType.AboveAverage => GetAboveAverageType(
							topNode,
							nameSpaceManager),
		  ExcelConditionalFormattingConstants.RuleType.Top10 => GetTop10Type(
							topNode,
							nameSpaceManager),
		  ExcelConditionalFormattingConstants.RuleType.TimePeriod => GetTimePeriodType(
							topNode,
							nameSpaceManager),
		  ExcelConditionalFormattingConstants.RuleType.CellIs => GetCellIs((XmlElement)topNode),
		  ExcelConditionalFormattingConstants.RuleType.BeginsWith => eExcelConditionalFormattingRuleType.BeginsWith,
		  //case ExcelConditionalFormattingConstants.RuleType.Between:
		  //  return eExcelConditionalFormattingRuleType.Between;
		  ExcelConditionalFormattingConstants.RuleType.ContainsBlanks => eExcelConditionalFormattingRuleType.ContainsBlanks,
		  ExcelConditionalFormattingConstants.RuleType.ContainsErrors => eExcelConditionalFormattingRuleType.ContainsErrors,
		  ExcelConditionalFormattingConstants.RuleType.ContainsText => eExcelConditionalFormattingRuleType.ContainsText,
		  ExcelConditionalFormattingConstants.RuleType.DuplicateValues => eExcelConditionalFormattingRuleType.DuplicateValues,
		  ExcelConditionalFormattingConstants.RuleType.EndsWith => eExcelConditionalFormattingRuleType.EndsWith,
		  //case ExcelConditionalFormattingConstants.RuleType.Equal:
		  //  return eExcelConditionalFormattingRuleType.Equal;
		  ExcelConditionalFormattingConstants.RuleType.Expression => eExcelConditionalFormattingRuleType.Expression,
		  //case ExcelConditionalFormattingConstants.RuleType.GreaterThan:
		  //  return eExcelConditionalFormattingRuleType.GreaterThan;
		  //case ExcelConditionalFormattingConstants.RuleType.GreaterThanOrEqual:
		  //  return eExcelConditionalFormattingRuleType.GreaterThanOrEqual;
		  //case ExcelConditionalFormattingConstants.RuleType.LessThan:
		  //  return eExcelConditionalFormattingRuleType.LessThan;
		  //case ExcelConditionalFormattingConstants.RuleType.LessThanOrEqual:
		  //  return eExcelConditionalFormattingRuleType.LessThanOrEqual;
		  //case ExcelConditionalFormattingConstants.RuleType.NotBetween:
		  //  return eExcelConditionalFormattingRuleType.NotBetween;
		  ExcelConditionalFormattingConstants.RuleType.NotContainsBlanks => eExcelConditionalFormattingRuleType.NotContainsBlanks,
		  ExcelConditionalFormattingConstants.RuleType.NotContainsErrors => eExcelConditionalFormattingRuleType.NotContainsErrors,
		  ExcelConditionalFormattingConstants.RuleType.NotContainsText => eExcelConditionalFormattingRuleType.NotContainsText,
		  //case ExcelConditionalFormattingConstants.RuleType.NotEqual:
		  //  return eExcelConditionalFormattingRuleType.NotEqual;
		  ExcelConditionalFormattingConstants.RuleType.UniqueValues => eExcelConditionalFormattingRuleType.UniqueValues,
		  ExcelConditionalFormattingConstants.RuleType.ColorScale => GetColorScaleType(
							topNode,
							nameSpaceManager),
		  ExcelConditionalFormattingConstants.RuleType.IconSet => GetIconSetType(topNode, nameSpaceManager),
		  ExcelConditionalFormattingConstants.RuleType.DataBar => eExcelConditionalFormattingRuleType.DataBar,
		  _ => throw new Exception(
					ExcelConditionalFormattingConstants.Errors.UnexpectedRuleTypeAttribute),
	  };

	private static eExcelConditionalFormattingRuleType GetCellIs(XmlElement node) => node.GetAttribute("operator") switch
	{
		ExcelConditionalFormattingConstants.Operators.BeginsWith => eExcelConditionalFormattingRuleType.BeginsWith,
		ExcelConditionalFormattingConstants.Operators.Between => eExcelConditionalFormattingRuleType.Between,
		ExcelConditionalFormattingConstants.Operators.ContainsText => eExcelConditionalFormattingRuleType.ContainsText,
		ExcelConditionalFormattingConstants.Operators.EndsWith => eExcelConditionalFormattingRuleType.EndsWith,
		ExcelConditionalFormattingConstants.Operators.Equal => eExcelConditionalFormattingRuleType.Equal,
		ExcelConditionalFormattingConstants.Operators.GreaterThan => eExcelConditionalFormattingRuleType.GreaterThan,
		ExcelConditionalFormattingConstants.Operators.GreaterThanOrEqual => eExcelConditionalFormattingRuleType.GreaterThanOrEqual,
		ExcelConditionalFormattingConstants.Operators.LessThan => eExcelConditionalFormattingRuleType.LessThan,
		ExcelConditionalFormattingConstants.Operators.LessThanOrEqual => eExcelConditionalFormattingRuleType.LessThanOrEqual,
		ExcelConditionalFormattingConstants.Operators.NotBetween => eExcelConditionalFormattingRuleType.NotBetween,
		ExcelConditionalFormattingConstants.Operators.NotContains => eExcelConditionalFormattingRuleType.NotContains,
		ExcelConditionalFormattingConstants.Operators.NotEqual => eExcelConditionalFormattingRuleType.NotEqual,
		_ => throw new Exception(
						  ExcelConditionalFormattingConstants.Errors.UnexistentOperatorTypeAttribute),
	};
	private static eExcelConditionalFormattingRuleType GetIconSetType(XmlNode topNode, XmlNamespaceManager nameSpaceManager)
	{
		var node = topNode.SelectSingleNode("d:iconSet/@iconSet", nameSpaceManager);
		if (node == null)
		{
			return eExcelConditionalFormattingRuleType.ThreeIconSet;
		}
		else
		{
			var v = node.Value;

			if (v[0] == '3')
			{
				return eExcelConditionalFormattingRuleType.ThreeIconSet;
			}
			else
			{
				return v[0] == '4' ? eExcelConditionalFormattingRuleType.FourIconSet : eExcelConditionalFormattingRuleType.FiveIconSet;
			}
		}
	}

	/// <summary>
	/// Get the "colorScale" rule type according to the number of "cfvo" and "color" nodes.
	/// If we have excatly 2 "cfvo" and "color" childs, then we return "twoColorScale"
	/// </summary>
	/// <returns>TwoColorScale or ThreeColorScale</returns>
	internal static eExcelConditionalFormattingRuleType GetColorScaleType(
	  XmlNode topNode,
	  XmlNamespaceManager nameSpaceManager)
	{
		// Get the <cfvo> nodes
		var cfvoNodes = topNode.SelectNodes(
		  string.Format(
			"{0}/{1}",
			ExcelConditionalFormattingConstants.Paths.ColorScale,
			ExcelConditionalFormattingConstants.Paths.Cfvo),
		  nameSpaceManager);

		// Get the <color> nodes
		var colorNodes = topNode.SelectNodes(
		  string.Format(
			"{0}/{1}",
			ExcelConditionalFormattingConstants.Paths.ColorScale,
			ExcelConditionalFormattingConstants.Paths.Color),
		  nameSpaceManager);

		// We determine if it is "TwoColorScale" or "ThreeColorScale" by the
		// number of <cfvo> and <color> inside the <colorScale> node
		if ((cfvoNodes == null) || (cfvoNodes.Count < 2) || (cfvoNodes.Count > 3)
		  || (colorNodes == null) || (colorNodes.Count < 2) || (colorNodes.Count > 3)
		  || (cfvoNodes.Count != colorNodes.Count))
		{
			throw new Exception(
			  ExcelConditionalFormattingConstants.Errors.WrongNumberCfvoColorNodes);
		}

		// Return the corresponding rule type (TwoColorScale or ThreeColorScale)
		return (cfvoNodes.Count == 2)
		  ? eExcelConditionalFormattingRuleType.TwoColorScale
		  : eExcelConditionalFormattingRuleType.ThreeColorScale;
	}

	/// <summary>
	/// Get the "aboveAverage" rule type according to the follwoing attributes:
	/// "AboveAverage", "EqualAverage" and "StdDev".
	/// 
	/// @StdDev greater than "0"                              == AboveStdDev
	/// @StdDev less than "0"                                 == BelowStdDev
	/// @AboveAverage = "1"/null and @EqualAverage = "0"/null == AboveAverage
	/// @AboveAverage = "1"/null and @EqualAverage = "1"      == AboveOrEqualAverage
	/// @AboveAverage = "0" and @EqualAverage = "0"/null      == BelowAverage
	/// @AboveAverage = "0" and @EqualAverage = "1"           == BelowOrEqualAverage
	/// /// </summary>
	/// <returns>AboveAverage, AboveOrEqualAverage, BelowAverage or BelowOrEqualAverage</returns>
	internal static eExcelConditionalFormattingRuleType GetAboveAverageType(
	  XmlNode topNode,
	  XmlNamespaceManager nameSpaceManager)
	{
		// Get @StdDev attribute
		var stdDev = ExcelConditionalFormattingHelper.GetAttributeIntNullable(
		  topNode,
		  ExcelConditionalFormattingConstants.Attributes.StdDev);

		if (stdDev > 0)
		{
			// @StdDev > "0" --> AboveStdDev
			return eExcelConditionalFormattingRuleType.AboveStdDev;
		}

		if (stdDev < 0)
		{
			// @StdDev < "0" --> BelowStdDev
			return eExcelConditionalFormattingRuleType.BelowStdDev;
		}

		// Get @AboveAverage attribute
		var isAboveAverage = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(
		  topNode,
		  ExcelConditionalFormattingConstants.Attributes.AboveAverage);

		// Get @EqualAverage attribute
		var isEqualAverage = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(
		  topNode,
		  ExcelConditionalFormattingConstants.Attributes.EqualAverage);

		if (isAboveAverage is null or true)
		{
			if (isEqualAverage == true)
			{
				// @AboveAverage = "1"/null and @EqualAverage = "1" == AboveOrEqualAverage
				return eExcelConditionalFormattingRuleType.AboveOrEqualAverage;
			}

			// @AboveAverage = "1"/null and @EqualAverage = "0"/null == AboveAverage
			return eExcelConditionalFormattingRuleType.AboveAverage;
		}

		if (isEqualAverage == true)
		{
			// @AboveAverage = "0" and @EqualAverage = "1" == BelowOrEqualAverage
			return eExcelConditionalFormattingRuleType.BelowOrEqualAverage;
		}

		// @AboveAverage = "0" and @EqualAverage = "0"/null == BelowAverage
		return eExcelConditionalFormattingRuleType.BelowAverage;
	}

	/// <summary>
	/// Get the "top10" rule type according to the follwoing attributes:
	/// "Bottom" and "Percent"
	/// 
	/// @Bottom = "1" and @Percent = "0"/null       == Bottom
	/// @Bottom = "1" and @Percent = "1"            == BottomPercent
	/// @Bottom = "0"/null and @Percent = "0"/null  == Top
	/// @Bottom = "0"/null and @Percent = "1"       == TopPercent
	/// /// </summary>
	/// <returns>Top, TopPercent, Bottom or BottomPercent</returns>
	public static eExcelConditionalFormattingRuleType GetTop10Type(
	  XmlNode topNode,
	  XmlNamespaceManager nameSpaceManager)
	{
		// Get @Bottom attribute
		var isBottom = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(
		  topNode,
		  ExcelConditionalFormattingConstants.Attributes.Bottom);

		// Get @Percent attribute
		var isPercent = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(
		  topNode,
		  ExcelConditionalFormattingConstants.Attributes.Percent);

		if (isBottom == true)
		{
			if (isPercent == true)
			{
				// @Bottom = "1" and @Percent = "1" == BottomPercent
				return eExcelConditionalFormattingRuleType.BottomPercent;
			}

			// @Bottom = "1" and @Percent = "0"/null == Bottom
			return eExcelConditionalFormattingRuleType.Bottom;
		}

		if (isPercent == true)
		{
			// @Bottom = "0"/null and @Percent = "1" == TopPercent
			return eExcelConditionalFormattingRuleType.TopPercent;
		}

		// @Bottom = "0"/null and @Percent = "0"/null == Top
		return eExcelConditionalFormattingRuleType.Top;
	}

	/// <summary>
	/// Get the "timePeriod" rule type according to "TimePeriod" attribute.
	/// /// </summary>
	/// <returns>Last7Days, LastMonth etc.</returns>
	public static eExcelConditionalFormattingRuleType GetTimePeriodType(
	  XmlNode topNode,
	  XmlNamespaceManager nameSpaceManager)
	{
		var timePeriod = ExcelConditionalFormattingTimePeriodType.GetTypeByAttribute(
		  ExcelConditionalFormattingHelper.GetAttributeString(
			topNode,
			ExcelConditionalFormattingConstants.Attributes.TimePeriod));

		return timePeriod switch
		{
			eExcelConditionalFormattingTimePeriodType.Last7Days => eExcelConditionalFormattingRuleType.Last7Days,
			eExcelConditionalFormattingTimePeriodType.LastMonth => eExcelConditionalFormattingRuleType.LastMonth,
			eExcelConditionalFormattingTimePeriodType.LastWeek => eExcelConditionalFormattingRuleType.LastWeek,
			eExcelConditionalFormattingTimePeriodType.NextMonth => eExcelConditionalFormattingRuleType.NextMonth,
			eExcelConditionalFormattingTimePeriodType.NextWeek => eExcelConditionalFormattingRuleType.NextWeek,
			eExcelConditionalFormattingTimePeriodType.ThisMonth => eExcelConditionalFormattingRuleType.ThisMonth,
			eExcelConditionalFormattingTimePeriodType.ThisWeek => eExcelConditionalFormattingRuleType.ThisWeek,
			eExcelConditionalFormattingTimePeriodType.Today => eExcelConditionalFormattingRuleType.Today,
			eExcelConditionalFormattingTimePeriodType.Tomorrow => eExcelConditionalFormattingRuleType.Tomorrow,
			eExcelConditionalFormattingTimePeriodType.Yesterday => eExcelConditionalFormattingRuleType.Yesterday,
			_ => throw new Exception(
					  ExcelConditionalFormattingConstants.Errors.UnexistentTimePeriodTypeAttribute),
		};
	}

	/// <summary>
	/// 
	/// </summary>
	/// <param name="type"></param>
	/// <returns></returns>
	public static string GetAttributeByType(
	  eExcelConditionalFormattingRuleType type) => type switch
	  {
		  eExcelConditionalFormattingRuleType.AboveAverage or eExcelConditionalFormattingRuleType.AboveOrEqualAverage or eExcelConditionalFormattingRuleType.BelowAverage or eExcelConditionalFormattingRuleType.BelowOrEqualAverage or eExcelConditionalFormattingRuleType.AboveStdDev or eExcelConditionalFormattingRuleType.BelowStdDev => ExcelConditionalFormattingConstants.RuleType.AboveAverage,
		  eExcelConditionalFormattingRuleType.Bottom or eExcelConditionalFormattingRuleType.BottomPercent or eExcelConditionalFormattingRuleType.Top or eExcelConditionalFormattingRuleType.TopPercent => ExcelConditionalFormattingConstants.RuleType.Top10,
		  eExcelConditionalFormattingRuleType.Last7Days or eExcelConditionalFormattingRuleType.LastMonth or eExcelConditionalFormattingRuleType.LastWeek or eExcelConditionalFormattingRuleType.NextMonth or eExcelConditionalFormattingRuleType.NextWeek or eExcelConditionalFormattingRuleType.ThisMonth or eExcelConditionalFormattingRuleType.ThisWeek or eExcelConditionalFormattingRuleType.Today or eExcelConditionalFormattingRuleType.Tomorrow or eExcelConditionalFormattingRuleType.Yesterday => ExcelConditionalFormattingConstants.RuleType.TimePeriod,
		  eExcelConditionalFormattingRuleType.Between or eExcelConditionalFormattingRuleType.Equal or eExcelConditionalFormattingRuleType.GreaterThan or eExcelConditionalFormattingRuleType.GreaterThanOrEqual or eExcelConditionalFormattingRuleType.LessThan or eExcelConditionalFormattingRuleType.LessThanOrEqual or eExcelConditionalFormattingRuleType.NotBetween or eExcelConditionalFormattingRuleType.NotEqual => ExcelConditionalFormattingConstants.RuleType.CellIs,
		  eExcelConditionalFormattingRuleType.ThreeIconSet or eExcelConditionalFormattingRuleType.FourIconSet or eExcelConditionalFormattingRuleType.FiveIconSet => ExcelConditionalFormattingConstants.RuleType.IconSet,
		  eExcelConditionalFormattingRuleType.ThreeColorScale or eExcelConditionalFormattingRuleType.TwoColorScale => ExcelConditionalFormattingConstants.RuleType.ColorScale,
		  eExcelConditionalFormattingRuleType.BeginsWith => ExcelConditionalFormattingConstants.RuleType.BeginsWith,
		  eExcelConditionalFormattingRuleType.ContainsBlanks => ExcelConditionalFormattingConstants.RuleType.ContainsBlanks,
		  eExcelConditionalFormattingRuleType.ContainsErrors => ExcelConditionalFormattingConstants.RuleType.ContainsErrors,
		  eExcelConditionalFormattingRuleType.ContainsText => ExcelConditionalFormattingConstants.RuleType.ContainsText,
		  eExcelConditionalFormattingRuleType.DuplicateValues => ExcelConditionalFormattingConstants.RuleType.DuplicateValues,
		  eExcelConditionalFormattingRuleType.EndsWith => ExcelConditionalFormattingConstants.RuleType.EndsWith,
		  eExcelConditionalFormattingRuleType.Expression => ExcelConditionalFormattingConstants.RuleType.Expression,
		  eExcelConditionalFormattingRuleType.NotContainsBlanks => ExcelConditionalFormattingConstants.RuleType.NotContainsBlanks,
		  eExcelConditionalFormattingRuleType.NotContainsErrors => ExcelConditionalFormattingConstants.RuleType.NotContainsErrors,
		  eExcelConditionalFormattingRuleType.NotContainsText => ExcelConditionalFormattingConstants.RuleType.NotContainsText,
		  eExcelConditionalFormattingRuleType.UniqueValues => ExcelConditionalFormattingConstants.RuleType.UniqueValues,
		  eExcelConditionalFormattingRuleType.DataBar => ExcelConditionalFormattingConstants.RuleType.DataBar,
		  _ => throw new Exception(
					ExcelConditionalFormattingConstants.Errors.MissingRuleType),
	  };

	/// <summary>
	/// Return cfvo §18.3.1.11 parent according to the rule type
	/// </summary>
	/// <param name="type"></param>
	/// <returns></returns>
	public static string GetCfvoParentPathByType(
	  eExcelConditionalFormattingRuleType type) => type switch
	  {
		  eExcelConditionalFormattingRuleType.TwoColorScale or eExcelConditionalFormattingRuleType.ThreeColorScale => ExcelConditionalFormattingConstants.Paths.ColorScale,
		  eExcelConditionalFormattingRuleType.ThreeIconSet or eExcelConditionalFormattingRuleType.FourIconSet or eExcelConditionalFormattingRuleType.FiveIconSet => ExcelConditionalFormattingConstants.RuleType.IconSet,
		  eExcelConditionalFormattingRuleType.DataBar => ExcelConditionalFormattingConstants.RuleType.DataBar,
		  _ => throw new Exception(
					ExcelConditionalFormattingConstants.Errors.MissingRuleType),
	  };
}