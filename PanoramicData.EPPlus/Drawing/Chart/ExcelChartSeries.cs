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
using System.Collections;
using System.Collections.Generic;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelBubbleChartSeries : ExcelChartSeries
{
	internal ExcelBubbleChartSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
	   : base(chart, ns, node, isPivot)
	{
	}
	public ExcelChartSerie Add(ExcelRangeBase Serie, ExcelRangeBase XSerie, ExcelRangeBase BubbleSize) => AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute, BubbleSize.FullAddressAbsolute);
	public ExcelChartSerie Add(string SerieAddress, string XSerieAddress, string BubbleSizeAddress) => AddSeries(SerieAddress, XSerieAddress, BubbleSizeAddress);
}
/// <summary>
/// Collection class for chart series
/// </summary>
public class ExcelChartSeries : XmlHelper, IEnumerable
{
	readonly List<ExcelChartSerie> _list = [];
	internal ExcelChart _chart;
	readonly XmlNode _node;
	readonly XmlNamespaceManager _ns;
	internal ExcelChartSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(ns, node)
	{
		_ns = ns;
		_chart = chart;
		_node = node;
		_isPivot = isPivot;
		SchemaNodeOrder = ["view3D", "plotArea", "barDir", "grouping", "scatterStyle", "varyColors", "ser", "explosion", "dLbls", "firstSliceAng", "holeSize", "shape", "legend", "axId"];
		foreach (XmlNode n in node.SelectNodes("c:ser", ns))
		{
			ExcelChartSerie s;
			if (chart.ChartNode.LocalName == "scatterChart")
			{
				s = new ExcelScatterChartSerie(this, ns, n, isPivot);
			}
			else if (chart.ChartNode.LocalName == "lineChart")
			{
				s = new ExcelLineChartSerie(this, ns, n, isPivot);
			}
			else if (chart.ChartNode.LocalName is "pieChart" or
					 "ofPieChart" or
					 "pie3DChart" or
					 "doughnutChart")
			{
				s = new ExcelPieChartSerie(this, ns, n, isPivot);
			}
			else if (chart.ChartNode.LocalName == "bubbleChart")
			{
				s = new ExcelBubbleChartSerie(this, ns, n, isPivot);
			}
			else
			{
				s = new ExcelChartSerie(this, ns, n, isPivot);
			}

			_list.Add(s);
		}
	}

	#region IEnumerable Members

	public IEnumerator GetEnumerator() => (_list.GetEnumerator());
	/// <summary>
	/// Returns the serie at the specified position.  
	/// </summary>
	/// <param name="PositionID">The position of the series.</param>
	/// <returns></returns>
	public ExcelChartSerie this[int PositionID]
	{
		get
		{
			return (_list[PositionID]);
		}
	}
	public int Count => _list.Count;
	/// <summary>
	/// Delete the chart at the specific position
	/// </summary>
	/// <param name="PositionID">Zero based</param>
	public void Delete(int PositionID)
	{
		var ser = _list[PositionID];
		ser.TopNode.ParentNode.RemoveChild(ser.TopNode);
		_list.RemoveAt(PositionID);
	}
	#endregion
	/// <summary>
	/// A reference to the chart object
	/// </summary>
	public ExcelChart Chart => _chart;
	#region "Add Series"

	/// <summary>
	/// Add a new serie to the chart. Do not apply to pivotcharts.
	/// </summary>
	/// <param name="Serie">The Y-Axis range</param>
	/// <param name="XSerie">The X-Axis range</param>
	/// <returns></returns>
	public virtual ExcelChartSerie Add(ExcelRangeBase Serie, ExcelRangeBase XSerie) => _chart.PivotTableSource != null
			? throw (new InvalidOperationException("Can't add a serie to a pivotchart"))
			: AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute, "");
	/// <summary>
	/// Add a new serie to the chart.Do not apply to pivotcharts.
	/// </summary>
	/// <param name="SerieAddress">The Y-Axis range</param>
	/// <param name="XSerieAddress">The X-Axis range</param>
	/// <returns></returns>
	public virtual ExcelChartSerie Add(string SerieAddress, string XSerieAddress) => _chart.PivotTableSource != null
			? throw (new InvalidOperationException("Can't add a serie to a pivotchart"))
			: AddSeries(SerieAddress, XSerieAddress, "");
	internal protected ExcelChartSerie AddSeries(string SeriesAddress, string XSeriesAddress, string bubbleSizeAddress)
	{
		var ser = _node.OwnerDocument.CreateElement("ser", ExcelPackage.schemaChart);
		var node = _node.SelectNodes("c:ser", _ns);
		if (node.Count > 0)
		{
			_node.InsertAfter(ser, node[^1]);
		}
		else
		{
			InserAfter(_node, "c:varyColors,c:grouping,c:barDir,c:scatterStyle,c:ofPieType", ser);
		}

		var idx = FindIndex();
		ser.InnerXml = string.Format("<c:idx val=\"{1}\" /><c:order val=\"{1}\" /><c:tx><c:strRef><c:f></c:f><c:strCache><c:ptCount val=\"1\" /></c:strCache></c:strRef></c:tx>{5}{0}{2}{3}{4}", AddExplosion(Chart.ChartType), idx, AddScatterPoint(Chart.ChartType), AddAxisNodes(Chart.ChartType), AddSmooth(Chart.ChartType), AddMarker(Chart.ChartType));
		var serie = CreateSerie(ser, SeriesAddress, XSeriesAddress, bubbleSizeAddress);
		serie.Series = SeriesAddress;
		serie.XSeries = XSeriesAddress;
		_list.Add(serie);
		return serie;
	}

	private ExcelChartSerie CreateSerie(XmlElement ser, string seriesAddress, string xSeriesAddress, string bubbleSizeAddress) => Chart.ChartType switch
	{
		eChartType.Bubble or
		eChartType.Bubble3DEffect => new ExcelBubbleChartSerie(this, NameSpaceManager, ser, _isPivot)
		{
			Bubble3D = Chart.ChartType == eChartType.Bubble3DEffect,
			Series = seriesAddress,
			XSeries = xSeriesAddress,
			BubbleSize = bubbleSizeAddress
		},
		eChartType.XYScatter or
		eChartType.XYScatterLines or
		eChartType.XYScatterLinesNoMarkers or
		eChartType.XYScatterSmooth or
		eChartType.XYScatterSmoothNoMarkers => new ExcelScatterChartSerie(this, NameSpaceManager, ser, _isPivot),
		eChartType.Radar or
		eChartType.RadarFilled or
		eChartType.RadarMarkers => new ExcelRadarChartSerie(this, NameSpaceManager, ser, _isPivot),
		eChartType.Surface or
		eChartType.SurfaceTopView or
		eChartType.SurfaceTopViewWireframe or
		eChartType.SurfaceWireframe => new ExcelSurfaceChartSerie(this, NameSpaceManager, ser, _isPivot),
		eChartType.Pie or
		eChartType.Pie3D or
		eChartType.PieExploded or
		eChartType.PieExploded3D or
		eChartType.PieOfPie or
		eChartType.Doughnut or
		eChartType.DoughnutExploded or
		eChartType.BarOfPie => new ExcelPieChartSerie(this, NameSpaceManager, ser, _isPivot),
		eChartType.Line or
		eChartType.LineMarkers or
		eChartType.LineMarkersStacked or
		eChartType.LineMarkersStacked100 or
		eChartType.LineStacked or
		eChartType.LineStacked100 => CreateLineSerie(ser),
		eChartType.BarClustered or
		eChartType.BarStacked or
		eChartType.BarStacked100 or
		eChartType.ColumnClustered or
		eChartType.ColumnStacked or
		eChartType.ColumnStacked100 or
		eChartType.BarClustered3D or
		eChartType.BarStacked3D or
		eChartType.BarStacked1003D or
		eChartType.ColumnClustered3D or
		eChartType.ColumnStacked3D or
		eChartType.ColumnStacked1003D or
		eChartType.ConeBarClustered or
		eChartType.ConeBarStacked or
		eChartType.ConeBarStacked100 or
		eChartType.ConeCol or
		eChartType.ConeColClustered or
		eChartType.ConeColStacked or
		eChartType.ConeColStacked100 or
		eChartType.CylinderBarClustered or
		eChartType.CylinderBarStacked or
		eChartType.CylinderBarStacked100 or
		eChartType.CylinderCol or
		eChartType.CylinderColClustered or
		eChartType.CylinderColStacked or
		eChartType.CylinderColStacked100 or
		eChartType.PyramidBarClustered or
		eChartType.PyramidBarStacked or
		eChartType.PyramidBarStacked100 or
		eChartType.PyramidCol or
		eChartType.PyramidColClustered or
		eChartType.PyramidColStacked or
		eChartType.PyramidColStacked100 => new ExcelBarChartSerie(this, NameSpaceManager, ser, _isPivot) { InvertIfNegative = false },
		_ => new ExcelChartSerie(this, NameSpaceManager, ser, _isPivot),
	};

	private ExcelLineChartSerie CreateLineSerie(XmlElement ser)
	{
		var serie = new ExcelLineChartSerie(this, NameSpaceManager, ser, _isPivot);
		if (Chart.ChartType is eChartType.LineMarkers or
			eChartType.LineMarkersStacked or
			eChartType.LineMarkersStacked100)
		{
			serie.Marker = eMarkerStyle.Square;
		}

		serie.Smooth = ((ExcelLineChart)Chart).Smooth;
		return serie;
	}
	bool _isPivot;
	internal void AddPivotSerie(ExcelPivotTable pivotTableSource)
	{
		var r = pivotTableSource.WorkSheet.Cells[pivotTableSource.Address.Address];
		_isPivot = true;
		AddSeries(r.Offset(0, 1, r._toRow - r._fromRow + 1, 1).FullAddressAbsolute, r.Offset(0, 0, r._toRow - r._fromRow + 1, 1).FullAddressAbsolute, "");
	}
	private int FindIndex()
	{
		int ret = 0, newID = 0;
		if (_chart.PlotArea.ChartTypes.Count > 1)
		{
			foreach (var chart in _chart.PlotArea.ChartTypes)
			{
				if (newID > 0)
				{
					foreach (ExcelChartSerie serie in chart.Series)
					{
						serie.SetID((++newID).ToString());
					}
				}
				else
				{
					if (chart == _chart)
					{
						ret += _list.Count + 1;
						newID = ret;
					}
					else
					{
						ret += chart.Series.Count;
					}
				}
			}

			return ret - 1;
		}
		else
		{
			return _list.Count;
		}
	}
	#endregion
	#region "Xml init Functions"
	private static string AddMarker(eChartType chartType) => chartType is eChartType.Line or
			eChartType.LineStacked or
			eChartType.LineStacked100 or
			eChartType.XYScatterLines or
			eChartType.XYScatterSmooth or
			eChartType.XYScatterLinesNoMarkers or
			eChartType.XYScatterSmoothNoMarkers
			? "<c:marker><c:symbol val=\"none\" /></c:marker>"
			: "";
	private static string AddScatterPoint(eChartType chartType) => chartType == eChartType.XYScatter ? "<c:spPr><a:ln w=\"28575\"><a:noFill /></a:ln></c:spPr>" : "";
	private static string AddAxisNodes(eChartType chartType) => chartType is eChartType.XYScatter or
			 eChartType.XYScatterLines or
			 eChartType.XYScatterLinesNoMarkers or
			 eChartType.XYScatterSmooth or
			 eChartType.XYScatterSmoothNoMarkers or
			 eChartType.Bubble or
			 eChartType.Bubble3DEffect
			? "<c:xVal /><c:yVal />"
			: "<c:val />";

	private static string AddExplosion(eChartType chartType) => chartType is eChartType.PieExploded3D or
		eChartType.PieExploded or
		eChartType.DoughnutExploded
		? "<c:explosion val=\"25\" />"
		: "";
	private static string AddSmooth(eChartType chartType) => chartType is eChartType.XYScatterSmooth or
		eChartType.XYScatterSmoothNoMarkers
		? "<c:smooth val=\"1\" />"
		: "";
	#endregion
}
