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
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Bar chart
/// </summary>
public sealed class ExcelBarChart : ExcelChart
{
	#region "Constructors"
	internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
		base(drawings, node, type, topChart, PivotTableSource)
	{
		SetTypeProperties(type);
	}

	internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
	   base(drawings, node, uriChart, part, chartXml, chartNode)
	{
	}

	internal ExcelBarChart(ExcelChart topChart, XmlNode chartNode) :
		base(topChart, chartNode)
	{
	}
	#endregion
	#region "Private functions"
	private void SetTypeProperties(eChartType type)
	{
		var direction = GetDirectionForType(type);
		if (direction.HasValue)
		{
			Direction = direction.Value;
		}

		var shape = GetShapeForType(type);
		if (shape.HasValue)
		{
			Shape = shape.Value;
		}
	}

	private static eDirection? GetDirectionForType(eChartType type)
	{
		if (type is eChartType.BarClustered or
			eChartType.BarStacked or
			eChartType.BarStacked100 or
			eChartType.BarClustered3D or
			eChartType.BarStacked3D or
			eChartType.BarStacked1003D or
			eChartType.ConeBarClustered or
			eChartType.ConeBarStacked or
			eChartType.ConeBarStacked100 or
			eChartType.CylinderBarClustered or
			eChartType.CylinderBarStacked or
			eChartType.CylinderBarStacked100 or
			eChartType.PyramidBarClustered or
			eChartType.PyramidBarStacked or
			eChartType.PyramidBarStacked100)
		{
			return eDirection.Bar;
		}

		return type is eChartType.ColumnClustered or
			eChartType.ColumnStacked or
			eChartType.ColumnStacked100 or
			eChartType.Column3D or
			eChartType.ColumnClustered3D or
			eChartType.ColumnStacked3D or
			eChartType.ColumnStacked1003D or
			eChartType.ConeCol or
			eChartType.ConeColClustered or
			eChartType.ConeColStacked or
			eChartType.ConeColStacked100 or
			eChartType.CylinderCol or
			eChartType.CylinderColClustered or
			eChartType.CylinderColStacked or
			eChartType.CylinderColStacked100 or
			eChartType.PyramidCol or
			eChartType.PyramidColClustered or
			eChartType.PyramidColStacked or
			eChartType.PyramidColStacked100
			? eDirection.Column
			: null;
	}

	private static eShape? GetShapeForType(eChartType type) => type switch
	{
		eChartType.Column3D or
		eChartType.ColumnClustered3D or
		eChartType.ColumnStacked3D or
		eChartType.ColumnStacked1003D or
		eChartType.BarClustered3D or
		eChartType.BarStacked3D or
		eChartType.BarStacked1003D => eShape.Box,
		eChartType.CylinderBarClustered or
		eChartType.CylinderBarStacked or
		eChartType.CylinderBarStacked100 or
		eChartType.CylinderCol or
		eChartType.CylinderColClustered or
		eChartType.CylinderColStacked or
		eChartType.CylinderColStacked100 => eShape.Cylinder,
		eChartType.ConeBarClustered or
		eChartType.ConeBarStacked or
		eChartType.ConeBarStacked100 or
		eChartType.ConeCol or
		eChartType.ConeColClustered or
		eChartType.ConeColStacked or
		eChartType.ConeColStacked100 => eShape.Cone,
		eChartType.PyramidBarClustered or
		eChartType.PyramidBarStacked or
		eChartType.PyramidBarStacked100 or
		eChartType.PyramidCol or
		eChartType.PyramidColClustered or
		eChartType.PyramidColStacked or
		eChartType.PyramidColStacked100 => eShape.Pyramid,
		_ => null,
	};
	#endregion
	#region "Properties"
	readonly string _directionPath = "c:barDir/@val";
	/// <summary>
	/// Direction, Bar or columns
	/// </summary>
	public eDirection Direction
	{
		get
		{
			return GetDirectionEnum(_chartXmlHelper.GetXmlNodeString(_directionPath));
		}
		internal set
		{
			_chartXmlHelper.SetXmlNodeString(_directionPath, GetDirectionText(value));
		}
	}
	readonly string _shapePath = "c:shape/@val";
	/// <summary>
	/// The shape of the bar/columns
	/// </summary>
	public eShape Shape
	{
		get
		{
			return GetShapeEnum(_chartXmlHelper.GetXmlNodeString(_shapePath));
		}
		internal set
		{
			_chartXmlHelper.SetXmlNodeString(_shapePath, GetShapeText(value));
		}
	}
	ExcelChartDataLabel? _dataLabel = null;
	/// <summary>
	/// Access to datalabel properties
	/// </summary>
	public ExcelChartDataLabel DataLabel => _dataLabel ??= new ExcelChartDataLabel(NameSpaceManager, ChartNode);
	readonly string _gapWidthPath = "c:gapWidth/@val";
	/// <summary>
	/// The size of the gap between two adjacent bars/columns
	/// </summary>
	public int GapWidth
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeInt(_gapWidthPath);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeString(_gapWidthPath, value.ToString(CultureInfo.InvariantCulture));
		}
	}
	#endregion
	#region "Direction Enum Traslation"
	private static string GetDirectionText(eDirection direction) => direction == eDirection.Bar ? "bar" : "col";

	private static eDirection GetDirectionEnum(string direction) => direction == "bar" ? eDirection.Bar : eDirection.Column;
	#endregion
	#region "Shape Enum Translation"
	private static string GetShapeText(eShape Shape) => Shape switch
	{
		eShape.Box => "box",
		eShape.Cone => "cone",
		eShape.ConeToMax => "coneToMax",
		eShape.Cylinder => "cylinder",
		eShape.Pyramid => "pyramid",
		eShape.PyramidToMax => "pyramidToMax",
		_ => "box",
	};

	private static eShape GetShapeEnum(string text) => text switch
	{
		"box" => eShape.Box,
		"cone" => eShape.Cone,
		"coneToMax" => eShape.ConeToMax,
		"cylinder" => eShape.Cylinder,
		"pyramid" => eShape.Pyramid,
		"pyramidToMax" => eShape.PyramidToMax,
		_ => eShape.Box,
	};
	#endregion
	internal override eChartType GetChartType(string name)
	{
		var type = name switch
		{
			"barChart" => GetBarChartType(),
			"bar3DChart" => GetBar3DChartType(),
			_ => null,
		};
		return type ?? base.GetChartType(name);
	}

	private static eChartType? ByGrouping(eGrouping grouping, eChartType stacked, eChartType percentStacked, eChartType clustered, eChartType? otherwise) => grouping switch
	{
		eGrouping.Stacked => stacked,
		eGrouping.PercentStacked => percentStacked,
		eGrouping.Clustered => clustered,
		_ => otherwise,
	};

	private eChartType? GetBarChartType() => Direction == eDirection.Bar
		? ByGrouping(Grouping, eChartType.BarStacked, eChartType.BarStacked100, eChartType.BarClustered, eChartType.BarClustered)
		: ByGrouping(Grouping, eChartType.ColumnStacked, eChartType.ColumnStacked100, eChartType.ColumnClustered, eChartType.ColumnClustered);

	// Bar variants of cone, cylinder and pyramid with no grouping fall back to the base type lookup.
	private eChartType? GetBar3DChartType()
	{
		var isBar = Direction == eDirection.Bar;
		var grouping = Grouping;
		return Shape switch
		{
			eShape.Box => isBar
				? ByGrouping(grouping, eChartType.BarStacked3D, eChartType.BarStacked1003D, eChartType.BarClustered3D, eChartType.BarClustered3D)
				: ByGrouping(grouping, eChartType.ColumnStacked3D, eChartType.ColumnStacked1003D, eChartType.ColumnClustered3D, eChartType.ColumnClustered3D),
			eShape.Cone or eShape.ConeToMax => isBar
				? ByGrouping(grouping, eChartType.ConeBarStacked, eChartType.ConeBarStacked100, eChartType.ConeBarClustered, null)
				: ByGrouping(grouping, eChartType.ConeColStacked, eChartType.ConeColStacked100, eChartType.ConeColClustered, eChartType.ConeCol),
			eShape.Cylinder => isBar
				? ByGrouping(grouping, eChartType.CylinderBarStacked, eChartType.CylinderBarStacked100, eChartType.CylinderBarClustered, null)
				: ByGrouping(grouping, eChartType.CylinderColStacked, eChartType.CylinderColStacked100, eChartType.CylinderColClustered, eChartType.CylinderCol),
			eShape.Pyramid or eShape.PyramidToMax => isBar
				? ByGrouping(grouping, eChartType.PyramidBarStacked, eChartType.PyramidBarStacked100, eChartType.PyramidBarClustered, null)
				: ByGrouping(grouping, eChartType.PyramidColStacked, eChartType.PyramidColStacked100, eChartType.PyramidColClustered, eChartType.PyramidCol),
			_ => null,
		};
	}
}