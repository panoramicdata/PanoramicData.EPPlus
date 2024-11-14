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
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
namespace OfficeOpenXml.Style;

/// <summary>
/// The numberformat of the cell
/// </summary>
public sealed class ExcelNumberFormat : StyleBase
{
	internal ExcelNumberFormat(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address, int index) :
		base(styles, ChangedEvent, PositionID, Address)
	{
		Index = index;
	}
	/// <summary>
	/// The numeric index fror the format
	/// </summary>
	public int NumFmtID => Index;
	/// <summary>
	/// The numberformat 
	/// </summary>
	public string Format
	{
		get
		{
			for (var i = 0; i < _styles.NumberFormats.Count; i++)
			{
				if (Index == _styles.NumberFormats[i].NumFmtId)
				{
					return _styles.NumberFormats[i].Format;
				}
			}

			return "general";
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Numberformat, eStyleProperty.Format, (string.IsNullOrEmpty(value) ? "General" : value), _positionID, _address));
		}
	}

	internal override string Id => Format;
	/// <summary>
	/// If the numeric format is a build-in from.
	/// </summary>
	public bool BuildIn { get; private set; }

	internal static string GetFromBuildInFromID(int _numFmtId) => _numFmtId switch
	{
		0 => "General",
		1 => "0",
		2 => "0.00",
		3 => "#,##0",
		4 => "#,##0.00",
		9 => "0%",
		10 => "0.00%",
		11 => "0.00E+00",
		12 => "# ?/?",
		13 => "# ??/??",
		14 => "mm-dd-yy",
		15 => "d-mmm-yy",
		16 => "d-mmm",
		17 => "mmm-yy",
		18 => "h:mm AM/PM",
		19 => "h:mm:ss AM/PM",
		20 => "h:mm",
		21 => "h:mm:ss",
		22 => "m/d/yy h:mm",
		37 => "#,##0 ;(#,##0)",
		38 => "#,##0 ;[Red](#,##0)",
		39 => "#,##0.00;(#,##0.00)",
		40 => "#,##0.00;[Red](#,##0.00)",
		45 => "mm:ss",
		46 => "[h]:mm:ss",
		47 => "mmss.0",
		48 => "##0.0",
		49 => "@",
		_ => string.Empty,
	};
	internal static int GetFromBuildIdFromFormat(string format) => format switch
	{
		"General" or "" => 0,
		"0" => 1,
		"0.00" => 2,
		"#,##0" => 3,
		"#,##0.00" => 4,
		"0%" => 9,
		"0.00%" => 10,
		"0.00E+00" => 11,
		"# ?/?" => 12,
		"# ??/??" => 13,
		"mm-dd-yy" => 14,
		"d-mmm-yy" => 15,
		"d-mmm" => 16,
		"mmm-yy" => 17,
		"h:mm AM/PM" => 18,
		"h:mm:ss AM/PM" => 19,
		"h:mm" => 20,
		"h:mm:ss" => 21,
		"m/d/yy h:mm" => 22,
		"#,##0 ;(#,##0)" => 37,
		"#,##0 ;[Red](#,##0)" => 38,
		"#,##0.00;(#,##0.00)" => 39,
		"#,##0.00;[Red](#,##0.00)" => 40,
		"mm:ss" => 45,
		"[h]:mm:ss" => 46,
		"mmss.0" => 47,
		"##0.0" => 48,
		"@" => 49,
		_ => int.MinValue,
	};
}
