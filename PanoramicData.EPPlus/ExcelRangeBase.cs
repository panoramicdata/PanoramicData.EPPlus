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
 * Jan Källman		    Initial Release		        2010-01-28
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 * Eyal Seagull		    Conditional Formatting      2012-04-03
 *******************************************************************************/
using OfficeOpenXml.Utils;
using System.Collections;

namespace OfficeOpenXml;

/// <summary>
/// A range of cells 
/// </summary>
public partial class ExcelRangeBase : ExcelAddress, IExcelCell, IDisposable, IEnumerable<ExcelRangeBase>, IEnumerator<ExcelRangeBase>
{
	/// <summary>
	/// Reference to the worksheet
	/// </summary>
	protected ExcelWorksheet _worksheet;
	internal ExcelWorkbook? _workbook;
	private delegate void _changeProp(ExcelRangeBase range, _setValue method, object? value);
	private delegate void _setValue(ExcelRangeBase range, object value, int row, int col);
	private _changeProp _changePropMethod;
	private int _styleID;
	private class CopiedCell
	{
		internal int Row { get; set; }
		internal int Column { get; set; }
		internal object? Value { get; set; }
		internal string Type { get; set; }
		internal object Formula { get; set; }
		internal int? StyleID { get; set; }
		internal Uri HyperLink { get; set; }
		internal ExcelComment Comment { get; set; }
		internal byte Flag { get; set; }
	}
	internal ExcelRangeBase(ExcelWorksheet xlWorksheet)
	{
		_worksheet = xlWorksheet;
		_ws = _worksheet.Name;
		_workbook = _worksheet.Workbook;
		SetDelegate();
	}
	/// <summary>
	/// On change address handler
	/// </summary>
	protected internal override void ChangeAddress()
	{
		if (Table != null)
		{
			SetRCFromTable(_workbook._package, null);
		}

		SetDelegate();
	}
	internal ExcelRangeBase(ExcelWorksheet xlWorksheet, string address) :
		base(xlWorksheet == null ? "" : xlWorksheet.Name, address)
	{
		_worksheet = xlWorksheet;
		_workbook = _worksheet.Workbook;
		SetRCFromTable(_worksheet._package, null);
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = _worksheet == null ? "" : _worksheet.Name;
		}

		SetDelegate();
	}
	internal ExcelRangeBase(ExcelWorkbook wb, ExcelWorksheet xlWorksheet, string address, bool isName) :
		base(xlWorksheet == null ? "" : xlWorksheet.Name, address, isName)
	{
		SetRCFromTable(wb._package, null);
		_worksheet = xlWorksheet;
		_workbook = wb;
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = (xlWorksheet?.Name);
		}

		SetDelegate();
	}
	private void SetToSelectedRange()
	{
		if (_worksheet.View.SelectedRange == "")
		{
			Address = "A1";
		}
		else
		{
			Address = _worksheet.View.SelectedRange;
		}
	}
	private void IsRangeValid(string type)
	{
		if (_fromRow <= 0)
		{
			if (_address == "")
			{
				SetToSelectedRange();
			}
			else
			{
				if (type == "")
				{
					throw new InvalidOperationException(string.Format("Range is not valid for this operation: {0}", _address));
				}
				else
				{
					throw new InvalidOperationException(string.Format("Range is not valid for {0} : {1}", type, _address));
				}
			}
		}
	}
	internal void UpdateAddress(string address) => throw new NotImplementedException();

	public void Dispose()
	{
	}

	CellsStoreEnumerator<ExcelCoreValue> _cellEnum;
	public IEnumerator<ExcelRangeBase> GetEnumerator()
	{
		Reset();
		return this;
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		Reset();
		return this;
	}

	/// <summary>
	/// The current range when enumerating
	/// </summary>
	public ExcelRangeBase Current => new(_worksheet, GetAddress(_cellEnum.Row, _cellEnum.Column));

	/// <summary>
	/// The current range when enumerating
	/// </summary>
	object IEnumerator.Current => ((object)(new ExcelRangeBase(_worksheet, GetAddress(_cellEnum.Row, _cellEnum.Column))));

	int _enumAddressIx = -1;
	public bool MoveNext()
	{
		if (_cellEnum.Next())
		{
			return true;
		}
		else if (_addresses != null)
		{
			_enumAddressIx++;
			if (_enumAddressIx < _addresses.Count)
			{
				_cellEnum = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values,
					_addresses[_enumAddressIx]._fromRow,
					_addresses[_enumAddressIx]._fromCol,
					_addresses[_enumAddressIx]._toRow,
					_addresses[_enumAddressIx]._toCol);
				return MoveNext();
			}
			else
			{
				return false;
			}
		}

		return false;
	}

	public void Reset()
	{
		_enumAddressIx = -1;
		_cellEnum = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
	}
}
