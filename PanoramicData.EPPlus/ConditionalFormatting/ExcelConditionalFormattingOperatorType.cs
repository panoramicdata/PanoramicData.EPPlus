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
/// Functions related to the <see cref="ExcelConditionalFormattingOperatorType"/>
/// </summary>
internal static class ExcelConditionalFormattingOperatorType
{
	/// <summary>
	/// 
	/// </summary>
	/// <param name="type"></param>
	/// <returns></returns>
	internal static string GetAttributeByType(
		eExcelConditionalFormattingOperatorType type) => type switch
		{
			eExcelConditionalFormattingOperatorType.BeginsWith => ExcelConditionalFormattingConstants.Operators.BeginsWith,
			eExcelConditionalFormattingOperatorType.Between => ExcelConditionalFormattingConstants.Operators.Between,
			eExcelConditionalFormattingOperatorType.ContainsText => ExcelConditionalFormattingConstants.Operators.ContainsText,
			eExcelConditionalFormattingOperatorType.EndsWith => ExcelConditionalFormattingConstants.Operators.EndsWith,
			eExcelConditionalFormattingOperatorType.Equal => ExcelConditionalFormattingConstants.Operators.Equal,
			eExcelConditionalFormattingOperatorType.GreaterThan => ExcelConditionalFormattingConstants.Operators.GreaterThan,
			eExcelConditionalFormattingOperatorType.GreaterThanOrEqual => ExcelConditionalFormattingConstants.Operators.GreaterThanOrEqual,
			eExcelConditionalFormattingOperatorType.LessThan => ExcelConditionalFormattingConstants.Operators.LessThan,
			eExcelConditionalFormattingOperatorType.LessThanOrEqual => ExcelConditionalFormattingConstants.Operators.LessThanOrEqual,
			eExcelConditionalFormattingOperatorType.NotBetween => ExcelConditionalFormattingConstants.Operators.NotBetween,
			eExcelConditionalFormattingOperatorType.NotContains => ExcelConditionalFormattingConstants.Operators.NotContains,
			eExcelConditionalFormattingOperatorType.NotEqual => ExcelConditionalFormattingConstants.Operators.NotEqual,
			_ => string.Empty,
		};

	/// <summary>
	/// 
	/// </summary>
	/// param name="attribute"
	/// <returns></returns>
	internal static eExcelConditionalFormattingOperatorType GetTypeByAttribute(
	  string attribute) => attribute switch
	  {
		  ExcelConditionalFormattingConstants.Operators.BeginsWith => eExcelConditionalFormattingOperatorType.BeginsWith,
		  ExcelConditionalFormattingConstants.Operators.Between => eExcelConditionalFormattingOperatorType.Between,
		  ExcelConditionalFormattingConstants.Operators.ContainsText => eExcelConditionalFormattingOperatorType.ContainsText,
		  ExcelConditionalFormattingConstants.Operators.EndsWith => eExcelConditionalFormattingOperatorType.EndsWith,
		  ExcelConditionalFormattingConstants.Operators.Equal => eExcelConditionalFormattingOperatorType.Equal,
		  ExcelConditionalFormattingConstants.Operators.GreaterThan => eExcelConditionalFormattingOperatorType.GreaterThan,
		  ExcelConditionalFormattingConstants.Operators.GreaterThanOrEqual => eExcelConditionalFormattingOperatorType.GreaterThanOrEqual,
		  ExcelConditionalFormattingConstants.Operators.LessThan => eExcelConditionalFormattingOperatorType.LessThan,
		  ExcelConditionalFormattingConstants.Operators.LessThanOrEqual => eExcelConditionalFormattingOperatorType.LessThanOrEqual,
		  ExcelConditionalFormattingConstants.Operators.NotBetween => eExcelConditionalFormattingOperatorType.NotBetween,
		  ExcelConditionalFormattingConstants.Operators.NotContains => eExcelConditionalFormattingOperatorType.NotContains,
		  ExcelConditionalFormattingConstants.Operators.NotEqual => eExcelConditionalFormattingOperatorType.NotEqual,
		  _ => throw new Exception(
					ExcelConditionalFormattingConstants.Errors.UnexistentOperatorTypeAttribute),
	  };
}