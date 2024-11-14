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
 * Mats Alm   		                Added       		        2011-03-23
 * Jan Källman		                License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using OfficeOpenXml.Utils;
using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation;

internal class RangeDataValidation : IRangeDataValidation
{
	public RangeDataValidation(ExcelWorksheet worksheet, string address)
	{
		Require.Argument(worksheet).IsNotNull("worksheet");
		Require.Argument(address).IsNotNullOrEmpty("address");
		_worksheet = worksheet;
		_address = address;
	}

	ExcelWorksheet _worksheet;
	string _address;

	public IExcelDataValidationAny AddAnyDataValidation() => _worksheet.DataValidations.AddAnyValidation(_address);

	public IExcelDataValidationInt AddIntegerDataValidation() => _worksheet.DataValidations.AddIntegerValidation(_address);

	public IExcelDataValidationDecimal AddDecimalDataValidation() => _worksheet.DataValidations.AddDecimalValidation(_address);

	public IExcelDataValidationDateTime AddDateTimeDataValidation() => _worksheet.DataValidations.AddDateTimeValidation(_address);

	public IExcelDataValidationList AddListDataValidation() => _worksheet.DataValidations.AddListValidation(_address);

	public IExcelDataValidationInt AddTextLengthDataValidation() => _worksheet.DataValidations.AddTextLengthValidation(_address);

	public IExcelDataValidationTime AddTimeDataValidation() => _worksheet.DataValidations.AddTimeValidation(_address);

	public IExcelDataValidationCustom AddCustomDataValidation() => _worksheet.DataValidations.AddCustomValidation(_address);
}
