using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.DataValidation;
using System.Text;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// WorkSheet object 
	/// </summary>
	public ExcelWorksheet Worksheet => _worksheet;
	/// <summary>
	/// Address including sheetname
	/// </summary>
	public new string FullAddress
	{
		get
		{
			if (Addresses?.Count > 0) //Contains multiple range.
			{
				var fullAddress = new StringBuilder(GetFullAddress(_worksheet.Name, Addresses[0].Address));
				for (var i = 1; i < Addresses.Count; i++)
				{
					fullAddress.Append(',').Append(GetFullAddress(_worksheet.Name, Addresses[i].Address));
				}

				return fullAddress.ToString();
			}

			return GetFullAddress(_worksheet.Name, _address);
		}
	}
	/// <summary>
	/// Address including sheetname
	/// </summary>
	public string FullAddressAbsolute
	{
		get
		{
			var wbwsRef = string.IsNullOrEmpty(_wb) ? _ws : "[" + _wb.Replace("'", "''") + "]" + _ws;
			if (Addresses == null)
			{
				return GetFullAddress(wbwsRef, GetAddress(_fromRow, _fromCol, _toRow, _toCol, true));
			}

			var fullAddress = new StringBuilder();
			foreach (var a in Addresses)
			{
				if (fullAddress.Length > 0)
				{
					fullAddress.Append(',');
				}

				if (a.Address == "#REF!")
				{
					fullAddress.Append(GetFullAddress(wbwsRef, "#REF!"));
				}
				else
				{
					fullAddress.Append(GetFullAddress(wbwsRef, GetAddress(a.Start.Row, a.Start.Column, a.End.Row, a.End.Column, true)));
				}
			}

			return fullAddress.ToString();
		}
	}
	/// <summary>
	/// Address including sheetname
	/// </summary>
	internal string FullAddressAbsoluteNoFullRowCol
	{
		get
		{
			var wbwsRef = string.IsNullOrEmpty(_wb) ? _ws : "[" + _wb.Replace("'", "''") + "]" + _ws;
			if (Addresses == null)
			{
				return GetFullAddress(wbwsRef, GetAddress(_fromRow, _fromCol, _toRow, _toCol, true), false);
			}

			var fullAddress = new StringBuilder();
			foreach (var a in Addresses)
			{
				if (fullAddress.Length > 0)
				{
					fullAddress.Append(',');
				}

				fullAddress.Append(GetFullAddress(wbwsRef, GetAddress(a.Start.Row, a.Start.Column, a.End.Row, a.End.Column, true), false));
			}

			return fullAddress.ToString();
		}
	}
	/// <summary>
	/// Conditional Formatting for this range.
	/// </summary>
	public IRangeConditionalFormatting ConditionalFormatting => new RangeConditionalFormatting(_worksheet, new ExcelAddress(Address));
	/// <summary>
	/// Data validation for this range.
	/// </summary>
	public IRangeDataValidation DataValidation => new RangeDataValidation(_worksheet, Address);
}
