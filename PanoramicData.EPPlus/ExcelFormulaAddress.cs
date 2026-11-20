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
using System.Text;

namespace OfficeOpenXml;

public class ExcelFormulaAddress : ExcelAddressBase
{
	internal ExcelFormulaAddress()
	{
	}

	public ExcelFormulaAddress(int fromRow, int fromCol, int toRow, int toColumn)
		: base(fromRow, fromCol, toRow, toColumn)
	{
		_ws = "";
	}
	public ExcelFormulaAddress(string address)
		: base(address)
	{
		SetFixed();
	}

	internal ExcelFormulaAddress(string ws, string address)
		: base(address)
	{
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ws;
		}

		SetFixed();
	}
	internal ExcelFormulaAddress(string ws, string address, bool isName)
		: base(address, isName)
	{
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ws;
		}

		if (!isName)
		{
			SetFixed();
		}
	}

	private void SetFixed()
	{
		if (Address.Contains('['))
		{
			return;
		}

		var address = FirstAddress;
		if (_fromRow == _toRow && _fromCol == _toCol)
		{
			GetFixed(address, out _fromRowFixed, out _fromColFixed);
		}
		else
		{
			var cells = address.Split(':');
			GetFixed(cells[0], out _fromRowFixed, out _fromColFixed);
			GetFixed(cells[1], out _toRowFixed, out _toColFixed);
		}
	}

	private static void GetFixed(string address, out bool rowFixed, out bool colFixed)
	{
		rowFixed = colFixed = false;
		var ix = address.IndexOf('$');
		while (ix > -1)
		{
			ix++;
			if (ix < address.Length)
			{
				if (address[ix] is >= '0' and <= '9')
				{
					rowFixed = true;
					break;
				}
				else
				{
					colFixed = true;
				}
			}

			ix = address.IndexOf('$', ix);
		}
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
				_address = GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _toRowFixed, _fromColFixed, _toColFixed);
			}

			return _address;
		}
		set
		{
			SetAddress(value);
			ChangeAddress();
			SetFixed();
		}
	}
	internal new List<ExcelFormulaAddress> _addresses;
	public new List<ExcelFormulaAddress> Addresses
	{
		get
		{
			_addresses ??= [];

			return _addresses;

		}
	}
	internal string GetOffset(int row, int column)
	{
		int fromRow = _fromRow, fromCol = _fromCol, toRow = _toRow, tocol = _toCol;
		var isMulti = (fromRow != toRow || fromCol != tocol);
		if (!_fromRowFixed)
		{
			fromRow += row;
		}

		if (!_fromColFixed)
		{
			fromCol += column;
		}

		if (isMulti)
		{
			if (!_toRowFixed)
			{
				toRow += row;
			}

			if (!_toColFixed)
			{
				tocol += column;
			}
		}
		else
		{
			toRow = fromRow;
			tocol = fromCol;
		}

		var a = new StringBuilder(GetAddress(fromRow, fromCol, toRow, tocol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed));
		if (Addresses != null)
		{
			foreach (var sa in Addresses)
			{
				a.Append(',').Append(sa.GetOffset(row, column));
			}
		}

		return a.ToString();
	}
}
