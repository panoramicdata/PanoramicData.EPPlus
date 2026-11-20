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
 *******************************************************************************
 * Jan Källman		Added		18-MAR-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
namespace OfficeOpenXml;

/// <summary>
/// Range address with the address property readonly
/// </summary>
public class ExcelAddress : ExcelAddressBase
{
	internal ExcelAddress()
	{
	}

	public ExcelAddress(int fromRow, int fromCol, int toRow, int toColumn)
		: base(fromRow, fromCol, toRow, toColumn)
	{
		_ws = "";
	}
	public ExcelAddress(string address)
		: base(address)
	{
	}

	internal ExcelAddress(string ws, string address)
		: base(address)
	{
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ws;
		}
	}
	internal ExcelAddress(string ws, string address, bool isName)
		: base(address, isName)
	{
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ws;
		}
	}

	public ExcelAddress(string Address, ExcelPackage package, ExcelAddressBase referenceAddress) :
		base(Address, package, referenceAddress)
	{

	}
	/// <summary>
	/// The address for the range
	/// </summary>
	/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
	public new string Address
	{
		get
		{
			if (string.IsNullOrEmpty(_address) && _fromRow > 0)
			{
				_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
			}

			return _address;
		}
		set
		{
			SetAddress(value);
			ChangeAddress();
		}
	}
}
