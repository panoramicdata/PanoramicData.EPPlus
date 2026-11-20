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
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using OfficeOpenXml.Style;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Position of the axis.
/// </summary>
public enum eAxisPosition
{
	Left = 0,
	Bottom = 1,
	Right = 2,
	Top = 3
}
/// <summary>
/// Position of the Y-Axis
/// </summary>
public enum eYAxisPosition
{
	Left = 0,
	Right = 2,
}
/// <summary>
/// Position of the X-Axis
/// </summary>
public enum eXAxisPosition
{
	Bottom = 1,
	Top = 3
}
/// <summary>
/// Axis orientaion
/// </summary>
public enum eAxisOrientation
{
	MaxMin,
	MinMax
}
/// <summary>
/// How the axis are crossed
/// </summary>
public enum eCrossBetween
{
	/// <summary>
	/// Specifies the value axis shall cross the category axis between data markers
	/// </summary>
	Between,
	/// <summary>
	/// Specifies the value axis shall cross the category axis at the midpoint of a category.
	/// </summary>
	MidCat
}
/// <summary>
/// Where the axis cross. 
/// </summary>
public enum eCrosses
{
	/// <summary>
	/// (Axis Crosses at Zero) The category axis crosses at the zero point of the valueaxis (if possible), or the minimum value (if theminimum is greater than zero) or the maximum (if the maximum is less than zero).
	/// </summary>
	AutoZero,
	/// <summary>
	/// The axis crosses at the maximum value
	/// </summary>
	Max,
	/// <summary>
	/// (Axis crosses at the minimum value of the chart.
	/// </summary>
	Min
}
/// <summary>
/// Tickmarks
/// </summary>
public enum eAxisTickMark
{
	/// <summary>
	/// Specifies the tick marks shall cross the axis. 
	/// </summary>
	Cross,
	/// <summary>
	/// Specifies the tick marks shall be inside the plot area. 
	/// </summary>
	In,
	/// <summary>
	/// Specifies there shall be no tick marks.
	/// </summary>
	None,
	/// <summary>
	/// Specifies the tick marks shall be outside the plot area.
	/// </summary>
	Out
}
public enum eBuildInUnits : long
{
	hundreds = 100,
	thousands = 1000,
	tenThousands = 10000,
	hundredThousands = 100000,
	millions = 1000000,
	tenMillions = 10000000,
	hundredMillions = 100000000,
	billions = 1000000000,
	trillions = 1000000000000
}
/// <summary>
/// An axis for a chart
/// </summary>
public sealed class ExcelChartAxis : XmlHelper
{
	/// <summary>
	/// Type of axis
	/// </summary>
	internal enum eAxisType
	{
		/// <summary>
		/// Value axis
		/// </summary>
		Val,
		/// <summary>
		/// Category axis
		/// </summary>
		Cat,
		/// <summary>
		/// Date axis
		/// </summary>
		Date,
		/// <summary>
		/// Series axis
		/// </summary>
		Serie
	}
	internal ExcelChartAxis(XmlNamespaceManager nameSpaceManager, XmlNode topNode) :
		base(nameSpaceManager, topNode)
	{
		SchemaNodeOrder = ["axId", "scaling", "logBase", "orientation", "max", "min", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crossesAt", "crosses", "crossBetween", "auto", "lblOffset", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit", "dispUnits", "spPr", "txPr"];
	}
	internal string Id => GetXmlNodeString("c:axId/@val");

	private T GetEnum<T>(string path, T defaultValue, bool ignoreCase) where T : struct, Enum
	{
		var v = GetXmlNodeString(path);
		return !string.IsNullOrEmpty(v) && Enum.TryParse(v, ignoreCase, out T result) ? result : defaultValue;
	}

	private void SetCamelCaseString(string path, Enum value)
	{
		var v = value.ToString();
		SetXmlNodeString(path, v[..1].ToLower(CultureInfo.InvariantCulture) + v[1..]);
	}

	private void SetDoubleOrDelete(string path, double? value)
	{
		if (value == null)
		{
			DeleteNode(path);
		}
		else
		{
			SetXmlNodeString(path, value.Value.ToString(CultureInfo.InvariantCulture));
		}
	}

	private double? GetUnit(string valuePath, string categoryPath) => GetXmlNodeDoubleNull(AxisType == eAxisType.Cat ? categoryPath : valuePath);

	private void SetUnit(string valuePath, string categoryPath, double? value)
	{
		if (value == null)
		{
			DeleteNode(valuePath);
			DeleteNode(categoryPath);
		}
		else
		{
			SetXmlNodeString(AxisType == eAxisType.Cat ? categoryPath : valuePath, value.Value.ToString(CultureInfo.InvariantCulture));
		}
	}

	private eTimeUnit? GetTimeUnit(string path) => GetXmlNodeString(path) switch
	{
		"years" => eTimeUnit.Years,
		"months" => eTimeUnit.Months,
		"days" => eTimeUnit.Days,
		_ => null,
	};

	private void SetTimeUnit(string path, eTimeUnit? value)
	{
		if (value.HasValue)
		{
			SetXmlNodeString(path, value.Value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
		else
		{
			DeleteNode(path);
		}
	}

	private ExcelDrawingBorder GetGridlines(string path)
	{
		if (TopNode.SelectSingleNode(path, NameSpaceManager) == null)
		{
			CreateNode(path);
		}

		return new ExcelDrawingBorder(NameSpaceManager, TopNode, $"{path}/c:spPr/a:ln");
	}

	const string _majorTickMark = "c:majorTickMark/@val";
	/// <summary>
	/// majorTickMark
	/// This element specifies the major tick marks for the axis.
	/// </summary>
	public eAxisTickMark MajorTickMark
	{
		get
		{
			return GetEnum(_majorTickMark, eAxisTickMark.Cross, false);
		}
		set
		{
			SetXmlNodeString(_majorTickMark, value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	const string _minorTickMark = "c:minorTickMark/@val";
	/// <summary>
	/// minorTickMark
	/// This element specifies the minor tick marks for the axis.
	/// </summary>
	public eAxisTickMark MinorTickMark
	{
		get
		{
			return GetEnum(_minorTickMark, eAxisTickMark.Cross, false);
		}
		set
		{
			SetXmlNodeString(_minorTickMark, value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}
	/// <summary>
	/// Type of axis
	/// </summary>
	internal eAxisType AxisType
	{
		get
		{
			try
			{
				return Enum.Parse<eAxisType>(TopNode.LocalName[..3], true);
			}
			catch
			{
				return eAxisType.Val;
			}
		}
	}
	private readonly string AXIS_POSITION_PATH = "c:axPos/@val";
	/// <summary>
	/// Where the axis is located
	/// </summary>
	public eAxisPosition AxisPosition
	{
		get
		{
			return GetXmlNodeString(AXIS_POSITION_PATH) switch
			{
				"b" => eAxisPosition.Bottom,
				"r" => eAxisPosition.Right,
				"t" => eAxisPosition.Top,
				_ => eAxisPosition.Left,
			};
		}
		internal set
		{
			SetXmlNodeString(AXIS_POSITION_PATH, value.ToString().ToLower(CultureInfo.InvariantCulture)[..1]);
		}
	}
	const string _crossesPath = "c:crosses/@val";
	/// <summary>
	/// Where the axis cross
	/// </summary>
	public eCrosses Crosses
	{
		get
		{
			return GetEnum(_crossesPath, eCrosses.AutoZero, true);
		}
		set
		{
			SetCamelCaseString(_crossesPath, value);
		}
	}
	const string _crossBetweenPath = "c:crossBetween/@val";
	/// <summary>
	/// How the axis are crossed
	/// </summary>
	public eCrossBetween CrossBetween
	{
		get
		{
			return GetEnum(_crossBetweenPath, eCrossBetween.Between, true);
		}
		set
		{
			SetCamelCaseString(_crossBetweenPath, value);
		}
	}
	const string _crossesAtPath = "c:crossesAt/@val";
	/// <summary>
	/// The value where the axis cross. 
	/// Null is automatic
	/// </summary>
	public double? CrossesAt
	{
		get
		{
			return GetXmlNodeDoubleNull(_crossesAtPath);
		}
		set
		{
			SetDoubleOrDelete(_crossesAtPath, value);
		}
	}
	const string _formatPath = "c:numFmt/@formatCode";
	/// <summary>
	/// Numberformat
	/// </summary>
	public string Format
	{
		get
		{
			return GetXmlNodeString(_formatPath);
		}
		set
		{
			SetXmlNodeString(_formatPath, value);
			SourceLinked = string.IsNullOrEmpty(value);
		}
	}
	const string _sourceLinkedPath = "c:numFmt/@sourceLinked";
	public bool SourceLinked
	{
		get
		{
			return GetXmlNodeBool(_sourceLinkedPath);
		}
		set
		{
			SetXmlNodeBool(_sourceLinkedPath, value);
		}
	}
	const string _lblPos = "c:tickLblPos/@val";
	/// <summary>
	/// Position of the labels
	/// </summary>
	public eTickLabelPosition LabelPosition
	{
		get
		{
			return GetEnum(_lblPos, eTickLabelPosition.NextTo, true);
		}
		set
		{
			SetCamelCaseString(_lblPos, value);
		}
	}
	ExcelDrawingFill _fill;
	/// <summary>
	/// Access to fill properties
	/// </summary>
	public ExcelDrawingFill Fill => ChartStyleAccess.GetFill(this, ref _fill);
	ExcelDrawingBorder _border;
	/// <summary>
	/// Access to border properties
	/// </summary>
	public ExcelDrawingBorder Border => ChartStyleAccess.GetBorder(this, ref _border);
	ExcelTextFont _font;
	/// <summary>
	/// Access to font properties
	/// </summary>
	public ExcelTextFont Font => ChartStyleAccess.GetFont(this, ref _font, ["pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t"]);
	/// <summary>
	/// If the axis is deleted
	/// </summary>
	public bool Deleted
	{
		get
		{
			return GetXmlNodeBool("c:delete/@val");
		}
		set
		{
			SetXmlNodeBool("c:delete/@val", value);
		}
	}
	const string _ticLblPos_Path = "c:tickLblPos/@val";
	/// <summary>
	/// Position of the Lables
	/// </summary>
	public eTickLabelPosition TickLabelPosition
	{
		get
		{
			var v = GetXmlNodeString(_ticLblPos_Path);
			return v == "" ? eTickLabelPosition.None : Enum.Parse<eTickLabelPosition>(v, true);
		}
		set
		{
			SetCamelCaseString(_ticLblPos_Path, value);
		}
	}
	const string _displayUnitPath = "c:dispUnits/c:builtInUnit/@val";
	const string _custUnitPath = "c:dispUnits/c:custUnit/@val";
	[SuppressMessage("Major Bug", "S1244:Floating point numbers should not be tested for equality", Justification = "Built-in units are exact whole numbers; a tolerance would change which unit is written.")]
	public double DisplayUnit
	{
		get
		{
			var v = GetXmlNodeString(_displayUnitPath);
			if (string.IsNullOrEmpty(v))
			{
				var c = GetXmlNodeDoubleNull(_custUnitPath);
				return c == null ? 0 : c.Value;
			}
			else
			{
				try
				{
					return (double)(long)Enum.Parse<eBuildInUnits>(v, true);
				}
				catch
				{
					return 0;
				}
			}
		}
		set
		{
			if (AxisType == eAxisType.Val && value >= 0)
			{
				foreach (var v in Enum.GetValues<eBuildInUnits>())
				{
					if ((double)(long)v == value)
					{
						DeleteNode(_custUnitPath);
						SetXmlNodeString(_displayUnitPath, ((eBuildInUnits)value).ToString());
						return;
					}
				}

				DeleteNode(_displayUnitPath);
				if (value != 0)
				{
					SetXmlNodeString(_custUnitPath, value.ToString(CultureInfo.InvariantCulture));
				}
			}
		}
	}
	ExcelChartTitle _title;
	/// <summary>
	/// Chart axis title
	/// </summary>
	public ExcelChartTitle Title
	{
		get
		{
			if (_title == null)
			{
				var node = TopNode.SelectSingleNode("c:title", NameSpaceManager);
				if (node == null)
				{
					CreateNode("c:title");
					node = TopNode.SelectSingleNode("c:title", NameSpaceManager);
					node.InnerXml = "<c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:r><a:t /></a:r></a:p></c:rich></c:tx><c:layout /><c:overlay val=\"0\" />";
				}

				_title = new ExcelChartTitle(NameSpaceManager, TopNode);
			}

			return _title;
		}
	}
	#region "Scaling"
	const string _minValuePath = "c:scaling/c:min/@val";
	/// <summary>
	/// Minimum value for the axis.
	/// Null is automatic
	/// </summary>
	public double? MinValue
	{
		get
		{
			return GetXmlNodeDoubleNull(_minValuePath);
		}
		set
		{
			SetDoubleOrDelete(_minValuePath, value);
		}
	}
	const string _maxValuePath = "c:scaling/c:max/@val";
	/// <summary>
	/// Max value for the axis.
	/// Null is automatic
	/// </summary>
	public double? MaxValue
	{
		get
		{
			return GetXmlNodeDoubleNull(_maxValuePath);
		}
		set
		{
			SetDoubleOrDelete(_maxValuePath, value);
		}
	}
	const string _majorUnitPath = "c:majorUnit/@val";
	const string _majorUnitCatPath = "c:tickLblSkip/@val";
	/// <summary>
	/// Major unit for the axis.
	/// Null is automatic
	/// </summary>
	public double? MajorUnit
	{
		get
		{
			return GetUnit(_majorUnitPath, _majorUnitCatPath);
		}
		set
		{
			SetUnit(_majorUnitPath, _majorUnitCatPath, value);
		}
	}
	const string _majorTimeUnitPath = "c:majorTimeUnit/@val";
	/// <summary>
	/// Major time unit for the axis.
	/// Null is automatic
	/// </summary>
	public eTimeUnit? MajorTimeUnit
	{
		get
		{
			return GetTimeUnit(_majorTimeUnitPath);
		}
		set
		{
			SetTimeUnit(_majorTimeUnitPath, value);
		}
	}
	const string _minorUnitPath = "c:minorUnit/@val";
	const string _minorUnitCatPath = "c:tickMarkSkip/@val";
	/// <summary>
	/// Minor unit for the axis.
	/// Null is automatic
	/// </summary>
	public double? MinorUnit
	{
		get
		{
			return GetUnit(_minorUnitPath, _minorUnitCatPath);
		}
		set
		{
			SetUnit(_minorUnitPath, _minorUnitCatPath, value);
		}
	}
	const string _minorTimeUnitPath = "c:minorTimeUnit/@val";
	/// <summary>
	/// Minor time unit for the axis.
	/// Null is automatic
	/// </summary>
	public eTimeUnit? MinorTimeUnit
	{
		get
		{
			return GetTimeUnit(_minorTimeUnitPath);
		}
		set
		{
			SetTimeUnit(_minorTimeUnitPath, value);
		}
	}
	const string _logbasePath = "c:scaling/c:logBase/@val";
	/// <summary>
	/// The base for a logaritmic scale
	/// Null for a normal scale
	/// </summary>
	public double? LogBase
	{
		get
		{
			return GetXmlNodeDoubleNull(_logbasePath);
		}
		set
		{
			if (value == null)
			{
				DeleteNode(_logbasePath);
			}
			else
			{
				var v = ((double)value);
				if (v is < 2 or > 1000)
				{
					throw (new ArgumentOutOfRangeException("Value must be between 2 and 1000"));
				}

				SetXmlNodeString(_logbasePath, v.ToString("0.0", CultureInfo.InvariantCulture));
			}
		}
	}
	const string _orientationPath = "c:scaling/c:orientation/@val";
	/// <summary>
	/// Axis orientation
	/// </summary>
	public eAxisOrientation Orientation
	{
		get
		{
			var v = GetXmlNodeString(_orientationPath);
			return v == "" ? eAxisOrientation.MinMax : Enum.Parse<eAxisOrientation>(v, true);
		}
		set
		{
			SetCamelCaseString(_orientationPath, value);
		}
	}
	#endregion

	#region GridLines 
	//Pull request from aesalazar
	const string _majorGridlinesPath = "c:majorGridlines";
	ExcelDrawingBorder _majorGridlines;

	/// <summary> 
	        /// Major Gridlines for the Axis 
	        /// </summary> 
	public ExcelDrawingBorder MajorGridlines => _majorGridlines ??= GetGridlines(_majorGridlinesPath);

	const string _minorGridlinesPath = "c:minorGridlines";
	ExcelDrawingBorder _minorGridlines;

	/// <summary> 
	        /// Minor Gridlines for the Axis 
	        /// </summary> 
	public ExcelDrawingBorder MinorGridlines => _minorGridlines ??= GetGridlines(_minorGridlinesPath);
	/// <summary> 
	        /// Removes Major and Minor gridlines from the Axis 
	        /// </summary> 
	public void RemoveGridlines() => RemoveGridlines(true, true);
	/// <summary>
	///  Removes gridlines from the Axis
	/// </summary>
	/// <param name="removeMajor">Indicates if the Major gridlines should be removed</param>
	/// <param name="removeMinor">Indicates if the Minor gridlines should be removed</param>
	public void RemoveGridlines(bool removeMajor, bool removeMinor)
	{
		if (removeMajor)
		{
			DeleteNode(_majorGridlinesPath);
			_majorGridlines = null;
		}

		if (removeMinor)
		{
			DeleteNode(_minorGridlinesPath);
			_minorGridlines = null;
		}
	}
	#endregion
}
