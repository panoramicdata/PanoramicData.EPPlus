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
 * Eyal Seagull    Conditional Formatting Adaption    2012-04-17
 *******************************************************************************/
using System;

namespace OfficeOpenXml.ConditionalFormatting;

/// <summary>
/// Functions related to the <see cref="ExcelConditionalFormattingTimePeriodType"/>
/// </summary>
internal static class ExcelConditionalFormattingTimePeriodType
{
	/// <summary>
	/// 
	/// </summary>
	/// <param name="type"></param>
	/// <returns></returns>
	public static string GetAttributeByType(
		eExcelConditionalFormattingTimePeriodType type) => type switch
		{
			eExcelConditionalFormattingTimePeriodType.Last7Days => ExcelConditionalFormattingConstants.TimePeriods.Last7Days,
			eExcelConditionalFormattingTimePeriodType.LastMonth => ExcelConditionalFormattingConstants.TimePeriods.LastMonth,
			eExcelConditionalFormattingTimePeriodType.LastWeek => ExcelConditionalFormattingConstants.TimePeriods.LastWeek,
			eExcelConditionalFormattingTimePeriodType.NextMonth => ExcelConditionalFormattingConstants.TimePeriods.NextMonth,
			eExcelConditionalFormattingTimePeriodType.NextWeek => ExcelConditionalFormattingConstants.TimePeriods.NextWeek,
			eExcelConditionalFormattingTimePeriodType.ThisMonth => ExcelConditionalFormattingConstants.TimePeriods.ThisMonth,
			eExcelConditionalFormattingTimePeriodType.ThisWeek => ExcelConditionalFormattingConstants.TimePeriods.ThisWeek,
			eExcelConditionalFormattingTimePeriodType.Today => ExcelConditionalFormattingConstants.TimePeriods.Today,
			eExcelConditionalFormattingTimePeriodType.Tomorrow => ExcelConditionalFormattingConstants.TimePeriods.Tomorrow,
			eExcelConditionalFormattingTimePeriodType.Yesterday => ExcelConditionalFormattingConstants.TimePeriods.Yesterday,
			_ => string.Empty,
		};

	/// <summary>
	/// 
	/// </summary>
	/// <param name="attribute"></param>
	/// <returns></returns>
	public static eExcelConditionalFormattingTimePeriodType GetTypeByAttribute(
	  string attribute) => attribute switch
	  {
		  ExcelConditionalFormattingConstants.TimePeriods.Last7Days => eExcelConditionalFormattingTimePeriodType.Last7Days,
		  ExcelConditionalFormattingConstants.TimePeriods.LastMonth => eExcelConditionalFormattingTimePeriodType.LastMonth,
		  ExcelConditionalFormattingConstants.TimePeriods.LastWeek => eExcelConditionalFormattingTimePeriodType.LastWeek,
		  ExcelConditionalFormattingConstants.TimePeriods.NextMonth => eExcelConditionalFormattingTimePeriodType.NextMonth,
		  ExcelConditionalFormattingConstants.TimePeriods.NextWeek => eExcelConditionalFormattingTimePeriodType.NextWeek,
		  ExcelConditionalFormattingConstants.TimePeriods.ThisMonth => eExcelConditionalFormattingTimePeriodType.ThisMonth,
		  ExcelConditionalFormattingConstants.TimePeriods.ThisWeek => eExcelConditionalFormattingTimePeriodType.ThisWeek,
		  ExcelConditionalFormattingConstants.TimePeriods.Today => eExcelConditionalFormattingTimePeriodType.Today,
		  ExcelConditionalFormattingConstants.TimePeriods.Tomorrow => eExcelConditionalFormattingTimePeriodType.Tomorrow,
		  ExcelConditionalFormattingConstants.TimePeriods.Yesterday => eExcelConditionalFormattingTimePeriodType.Yesterday,
		  _ => throw new Exception(
					ExcelConditionalFormattingConstants.Errors.UnexistentTimePeriodTypeAttribute),
	  };
}