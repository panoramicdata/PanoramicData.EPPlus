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
 * Mats Alm   		                Added       		        2011-01-01
 * Jan Källman		                License changed GPL-->LGPL  2011-12-27
 * Raziq York 		                Added support for Any type  2014-08-08
 *******************************************************************************/
using System;

namespace OfficeOpenXml.DataValidation;

/// <summary>
/// Enum for available data validation types
/// </summary>
public enum eDataValidationType
{
	/// <summary>
	/// Any value
	/// </summary>
	Any,
	/// <summary>
	/// Integer value
	/// </summary>
	Whole,
	/// <summary>
	/// Decimal values
	/// </summary>
	Decimal,
	/// <summary>
	/// List of values
	/// </summary>
	List,
	/// <summary>
	/// Text length validation
	/// </summary>
	TextLength,
	/// <summary>
	/// DateTime validation
	/// </summary>
	DateTime,
	/// <summary>
	/// Time validation
	/// </summary>
	Time,
	/// <summary>
	/// Custom validation
	/// </summary>
	Custom
}

internal static class DataValidationSchemaNames
{
	public const string Any = "";
	public const string Whole = "whole";
	public const string Decimal = "decimal";
	public const string List = "list";
	public const string TextLength = "textLength";
	public const string Date = "date";
	public const string Time = "time";
	public const string Custom = "custom";
}

/// <summary>
/// Types of datavalidation
/// </summary>
public class ExcelDataValidationType
{
	private ExcelDataValidationType(eDataValidationType validationType, bool allowOperator, string schemaName)
	{
		Type = validationType;
		AllowOperator = allowOperator;
		SchemaName = schemaName;
	}

	/// <summary>
	/// Validation type
	/// </summary>
	public eDataValidationType Type
	{
		get;
		private set;
	}

	internal string SchemaName
	{
		get;
		private set;
	}

	/// <summary>
	/// This type allows operator to be set
	/// </summary>
	internal bool AllowOperator
	{

		get;
		private set;
	}

	/// <summary>
	/// Returns a validation type by <see cref="eDataValidationType"/>
	/// </summary>
	/// <param name="type"></param>
	/// <returns></returns>
	internal static ExcelDataValidationType GetByValidationType(eDataValidationType type) => type switch
	{
		eDataValidationType.Any => ExcelDataValidationType.Any,
		eDataValidationType.Whole => ExcelDataValidationType.Whole,
		eDataValidationType.List => ExcelDataValidationType.List,
		eDataValidationType.Decimal => ExcelDataValidationType.Decimal,
		eDataValidationType.TextLength => ExcelDataValidationType.TextLength,
		eDataValidationType.DateTime => ExcelDataValidationType.DateTime,
		eDataValidationType.Time => ExcelDataValidationType.Time,
		eDataValidationType.Custom => ExcelDataValidationType.Custom,
		_ => throw new InvalidOperationException("Non supported Validationtype : " + type.ToString()),
	};

	internal static ExcelDataValidationType GetBySchemaName(string schemaName) => schemaName switch
	{
		DataValidationSchemaNames.Any => ExcelDataValidationType.Any,
		DataValidationSchemaNames.Whole => ExcelDataValidationType.Whole,
		DataValidationSchemaNames.Decimal => ExcelDataValidationType.Decimal,
		DataValidationSchemaNames.List => ExcelDataValidationType.List,
		DataValidationSchemaNames.TextLength => ExcelDataValidationType.TextLength,
		DataValidationSchemaNames.Date => ExcelDataValidationType.DateTime,
		DataValidationSchemaNames.Time => ExcelDataValidationType.Time,
		DataValidationSchemaNames.Custom => ExcelDataValidationType.Custom,
		_ => throw new ArgumentException("Invalid schemaname: " + schemaName),
	};

	/// <summary>
	/// Overridden Equals, compares on internal validation type
	/// </summary>
	/// <param name="obj"></param>
	/// <returns></returns>
	public override bool Equals(object obj) => obj is ExcelDataValidationType && ((ExcelDataValidationType)obj).Type == Type;

	/// <summary>
	/// Overrides GetHashCode()
	/// </summary>
	/// <returns></returns>
	public override int GetHashCode() => base.GetHashCode();

	/// <summary>
	/// Integer values
	/// </summary>
	private static ExcelDataValidationType _any;
	public static ExcelDataValidationType Any
	{
		get
		{
			_any ??= new ExcelDataValidationType(eDataValidationType.Any, false, DataValidationSchemaNames.Any);

			return _any;
		}
	}

	/// <summary>
	/// Integer values
	/// </summary>
	private static ExcelDataValidationType _whole;
	public static ExcelDataValidationType Whole
	{
		get
		{
			_whole ??= new ExcelDataValidationType(eDataValidationType.Whole, true, DataValidationSchemaNames.Whole);

			return _whole;
		}
	}

	/// <summary>
	/// List of allowed values
	/// </summary>
	private static ExcelDataValidationType _list;
	public static ExcelDataValidationType List
	{
		get
		{
			_list ??= new ExcelDataValidationType(eDataValidationType.List, false, DataValidationSchemaNames.List);

			return _list;
		}
	}

	private static ExcelDataValidationType _decimal;
	public static ExcelDataValidationType Decimal
	{
		get
		{
			_decimal ??= new ExcelDataValidationType(eDataValidationType.Decimal, true, DataValidationSchemaNames.Decimal);

			return _decimal;
		}
	}

	private static ExcelDataValidationType _textLength;
	public static ExcelDataValidationType TextLength
	{
		get
		{
			_textLength ??= new ExcelDataValidationType(eDataValidationType.TextLength, true, DataValidationSchemaNames.TextLength);

			return _textLength;
		}
	}

	private static ExcelDataValidationType _dateTime;
	public static ExcelDataValidationType DateTime
	{
		get
		{
			_dateTime ??= new ExcelDataValidationType(eDataValidationType.DateTime, true, DataValidationSchemaNames.Date);

			return _dateTime;
		}
	}

	private static ExcelDataValidationType _time;
	public static ExcelDataValidationType Time
	{
		get
		{
			_time ??= new ExcelDataValidationType(eDataValidationType.Time, true, DataValidationSchemaNames.Time);

			return _time;
		}
	}

	private static ExcelDataValidationType _custom;
	public static ExcelDataValidationType Custom
	{
		get
		{
			_custom ??= new ExcelDataValidationType(eDataValidationType.Custom, true, DataValidationSchemaNames.Custom);

			return _custom;
		}
	}
}
