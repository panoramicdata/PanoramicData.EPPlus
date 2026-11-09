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
/// A range address
/// </summary>
/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
public partial class ExcelAddressBase : ExcelCellBase
{
	internal protected int _fromRow = -1, _toRow, _fromCol, _toCol;
	protected internal bool _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed;
	internal protected string _wb;
	internal protected string _ws;
	internal protected string _address;

	internal enum eAddressCollition
	{
		No,
		Partly,
		Inside,
		Equal
	}
	internal enum eShiftType
	{
		Right,
		Down,
		EntireRow,
		EntireColumn
	}
	internal ExcelAddressBase()
	{
	}
	/// <summary>
	/// Creates an Address object
	/// </summary>
	/// <param name="fromRow">start row</param>
	/// <param name="fromCol">start column</param>
	/// <param name="toRow">End row</param>
	/// <param name="toColumn">End column</param>
	public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn)
	{
		_fromRow = fromRow;
		_toRow = toRow;
		_fromCol = fromCol;
		_toCol = toColumn;
		Validate();

		_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
	}
	/// <summary>
	/// Creates an Address object
	/// </summary>
	/// <param name="worksheetName">Worksheet Name</param>
	/// <param name="fromRow">start row</param>
	/// <param name="fromCol">start column</param>
	/// <param name="toRow">End row</param>
	/// <param name="toColumn">End column</param>
	public ExcelAddressBase(string worksheetName, int fromRow, int fromCol, int toRow, int toColumn)
	{
		_ws = worksheetName;
		_fromRow = fromRow;
		_toRow = toRow;
		_fromCol = fromCol;
		_toCol = toColumn;
		Validate();

		_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
	}
	/// <summary>
	/// Creates an Address object
	/// </summary>
	/// <param name="fromRow">start row</param>
	/// <param name="fromCol">start column</param>
	/// <param name="toRow">End row</param>
	/// <param name="toColumn">End column</param>
	/// <param name="fromRowFixed">start row fixed</param>
	/// <param name="fromColFixed">start column fixed</param>
	/// <param name="toRowFixed">End row fixed</param>
	/// <param name="toColFixed">End column fixed</param>
	public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn, bool fromRowFixed, bool fromColFixed, bool toRowFixed, bool toColFixed)
	{
		_fromRow = fromRow;
		_toRow = toRow;
		_fromCol = fromCol;
		_toCol = toColumn;
		_fromRowFixed = fromRowFixed;
		_fromColFixed = fromColFixed;
		_toRowFixed = toRowFixed;
		_toColFixed = toColFixed;
		Validate();

		_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, fromColFixed, _toRowFixed, _toColFixed);
	}
	/// <summary>
	/// Creates an Address object
	/// </summary>
	/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
	/// <param name="address">The Excel Address</param>
	public ExcelAddressBase(string address)
	{
		SetAddress(address);
	}
	/// <summary>
	/// Creates an Address object
	/// </summary>
	/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
	/// <param name="address">The Excel Address</param>
	/// <param name="pck">Reference to the package to find information about tables and names</param>
	/// <param name="referenceAddress">The address</param>
	public ExcelAddressBase(string address, ExcelPackage pck, ExcelAddressBase referenceAddress)
	{
		SetAddress(address);
		SetRCFromTable(pck, referenceAddress);
	}

	/// <summary>
	/// Address is an defined name
	/// </summary>
	/// <param name="address">the name</param>
	/// <param name="isName">Should always be true</param>
	internal ExcelAddressBase(string address, bool isName)
	{
		if (isName)
		{
			_address = address;
			_fromRow = -1;
			_fromCol = -1;
			_toRow = -1;
			_toCol = -1;
			_start = null;
			_end = null;
		}
		else
		{
			SetAddress(address);
		}
	}

	protected internal void SetAddress(string address)
	{
		address = address.Trim();
		if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(address, "'") || Utils.ConvertUtil._invariantCompareInfo.IsPrefix(address, "["))
		{
			SetWbWs(address);
		}
		else
		{
			_address = address;
		}

		_addresses = null;
		if (_address.IndexOfAny([',', '!', '[']) > -1)
		{
			//Advanced address. Including Sheet or multi or table.
			ExtractAddress(_address);
		}
		else
		{
			//Simple address
			GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
			_start = null;
			_end = null;
		}

		_address = address;
		Validate();
	}
	internal protected virtual void ChangeAddress()
	{
	}
}
