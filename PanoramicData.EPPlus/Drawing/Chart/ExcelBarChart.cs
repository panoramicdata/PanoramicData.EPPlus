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
	//internal ExcelBarChart(ExcelDrawings drawings, XmlNode node) :
	//    base(drawings, node/*, 1*/)
	//{
	//    SetChartNodeText("");
	//}
	//internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
	//    base(drawings, node, type)
	//{
	//    SetChartNodeText("");

	//    SetTypeProperties(drawings, type);
	//}
	internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
		base(drawings, node, type, topChart, PivotTableSource)
	{
		SetChartNodeText("");

		SetTypeProperties(drawings, type);
	}

	internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
	   base(drawings, node, uriChart, part, chartXml, chartNode)
	{
		SetChartNodeText(chartNode.Name);
	}

	internal ExcelBarChart(ExcelChart topChart, XmlNode chartNode) :
		base(topChart, chartNode)
	{
		SetChartNodeText(chartNode.Name);
	}
	#endregion
	#region "Private functions"
	//string _chartTopPath="c:chartSpace/c:chart/c:plotArea/{0}";
	private void SetChartNodeText(string chartNodeText)
	{
		if (string.IsNullOrEmpty(chartNodeText))
		{
			chartNodeText = GetChartNodeText();
		}
		//_chartTopPath = string.Format(_chartTopPath, chartNodeText);
		//_directionPath = string.Format(_directionPath, _chartTopPath);
		//_shapePath = string.Format(_shapePath, _chartTopPath);
	}
	private void SetTypeProperties(ExcelDrawings drawings, eChartType type)
	{
		/******* Bar direction *******/
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
			Direction = eDirection.Bar;
		}
		else if (
			type is eChartType.ColumnClustered or
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
			eChartType.PyramidColStacked100)
		{
			Direction = eDirection.Column;
		}

		/****** Shape ******/
		if (/*type == eChartType.ColumnClustered ||
                type == eChartType.ColumnStacked ||
                type == eChartType.ColumnStacked100 ||*/
			type is eChartType.Column3D or
			eChartType.ColumnClustered3D or
			eChartType.ColumnStacked3D or
			eChartType.ColumnStacked1003D or
			/*type == eChartType.BarClustered ||
                type == eChartType.BarStacked ||
                type == eChartType.BarStacked100 ||*/
			eChartType.BarClustered3D or
			eChartType.BarStacked3D or
			eChartType.BarStacked1003D)
		{
			Shape = eShape.Box;
		}
		else if (
			type is eChartType.CylinderBarClustered or
			eChartType.CylinderBarStacked or
			eChartType.CylinderBarStacked100 or
			eChartType.CylinderCol or
			eChartType.CylinderColClustered or
			eChartType.CylinderColStacked or
			eChartType.CylinderColStacked100)
		{
			Shape = eShape.Cylinder;
		}
		else if (
			type is eChartType.ConeBarClustered or
			eChartType.ConeBarStacked or
			eChartType.ConeBarStacked100 or
			eChartType.ConeCol or
			eChartType.ConeColClustered or
			eChartType.ConeColStacked or
			eChartType.ConeColStacked100)
		{
			Shape = eShape.Cone;
		}
		else if (
			type is eChartType.PyramidBarClustered or
			eChartType.PyramidBarStacked or
			eChartType.PyramidBarStacked100 or
			eChartType.PyramidCol or
			eChartType.PyramidColClustered or
			eChartType.PyramidColStacked or
			eChartType.PyramidColStacked100)
		{
			Shape = eShape.Pyramid;
		}
	}
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
	ExcelChartDataLabel _DataLabel = null;
	/// <summary>
	/// Access to datalabel properties
	/// </summary>
	public ExcelChartDataLabel DataLabel
	{
		get
		{
			_DataLabel ??= new ExcelChartDataLabel(NameSpaceManager, ChartNode);

			return _DataLabel;
		}
	}
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
	private static string GetDirectionText(eDirection direction) => direction switch
	{
		eDirection.Bar => "bar",
		_ => "col",
	};

	private static eDirection GetDirectionEnum(string direction) => direction switch
	{
		"bar" => eDirection.Bar,
		_ => eDirection.Column,
	};
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
		if (name == "barChart")
		{
			if (Direction == eDirection.Bar)
			{
				if (Grouping == eGrouping.Stacked)
				{
					return eChartType.BarStacked;
				}
				else
				{
					return Grouping == eGrouping.PercentStacked ? eChartType.BarStacked100 : eChartType.BarClustered;
				}
			}
			else
			{
				if (Grouping == eGrouping.Stacked)
				{
					return eChartType.ColumnStacked;
				}
				else
				{
					return Grouping == eGrouping.PercentStacked ? eChartType.ColumnStacked100 : eChartType.ColumnClustered;
				}
			}
		}

		if (name == "bar3DChart")
		{
			#region "Bar Shape"
			if (Shape == eShape.Box)
			{
				if (Direction == eDirection.Bar)
				{
					if (Grouping == eGrouping.Stacked)
					{
						return eChartType.BarStacked3D;
					}
					else
					{
						return Grouping == eGrouping.PercentStacked ? eChartType.BarStacked1003D : eChartType.BarClustered3D;
					}
				}
				else
				{
					if (Grouping == eGrouping.Stacked)
					{
						return eChartType.ColumnStacked3D;
					}
					else
					{
						return Grouping == eGrouping.PercentStacked ? eChartType.ColumnStacked1003D : eChartType.ColumnClustered3D;
					}
				}
			}
			#endregion
			#region "Cone Shape"
			if (Shape is eShape.Cone or eShape.ConeToMax)
			{
				if (Direction == eDirection.Bar)
				{
					if (Grouping == eGrouping.Stacked)
					{
						return eChartType.ConeBarStacked;
					}
					else if (Grouping == eGrouping.PercentStacked)
					{
						return eChartType.ConeBarStacked100;
					}
					else if (Grouping == eGrouping.Clustered)
					{
						return eChartType.ConeBarClustered;
					}
				}
				else
				{
					if (Grouping == eGrouping.Stacked)
					{
						return eChartType.ConeColStacked;
					}
					else if (Grouping == eGrouping.PercentStacked)
					{
						return eChartType.ConeColStacked100;
					}
					else
					{
						return Grouping == eGrouping.Clustered ? eChartType.ConeColClustered : eChartType.ConeCol;
					}
				}
			}
			#endregion
			#region "Cylinder Shape"
			if (Shape == eShape.Cylinder)
			{
				if (Direction == eDirection.Bar)
				{
					if (Grouping == eGrouping.Stacked)
					{
						return eChartType.CylinderBarStacked;
					}
					else if (Grouping == eGrouping.PercentStacked)
					{
						return eChartType.CylinderBarStacked100;
					}
					else if (Grouping == eGrouping.Clustered)
					{
						return eChartType.CylinderBarClustered;
					}
				}
				else
				{
					if (Grouping == eGrouping.Stacked)
					{
						return eChartType.CylinderColStacked;
					}
					else if (Grouping == eGrouping.PercentStacked)
					{
						return eChartType.CylinderColStacked100;
					}
					else
					{
						return Grouping == eGrouping.Clustered ? eChartType.CylinderColClustered : eChartType.CylinderCol;
					}
				}
			}
			#endregion
			#region "Pyramide Shape"
			if (Shape is eShape.Pyramid or eShape.PyramidToMax)
			{
				if (Direction == eDirection.Bar)
				{
					if (Grouping == eGrouping.Stacked)
					{
						return eChartType.PyramidBarStacked;
					}
					else if (Grouping == eGrouping.PercentStacked)
					{
						return eChartType.PyramidBarStacked100;
					}
					else if (Grouping == eGrouping.Clustered)
					{
						return eChartType.PyramidBarClustered;
					}
				}
				else
				{
					if (Grouping == eGrouping.Stacked)
					{
						return eChartType.PyramidColStacked;
					}
					else if (Grouping == eGrouping.PercentStacked)
					{
						return eChartType.PyramidColStacked100;
					}
					else
					{
						return Grouping == eGrouping.Clustered ? eChartType.PyramidColClustered : eChartType.PyramidCol;
					}
				}
			}
			#endregion
		}

		return base.GetChartType(name);
	}
}
