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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using System.IO;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging;
namespace OfficeOpenXml.Drawing.Chart;

#region "Chart Enums"
/// <summary>
/// Chart type
/// </summary>
public enum eChartType
{
	Area3D = -4098,
	AreaStacked3D = 78,
	AreaStacked1003D = 79,
	BarClustered3D = 60,
	BarStacked3D = 61,
	BarStacked1003D = 62,
	Column3D = -4100,
	ColumnClustered3D = 54,
	ColumnStacked3D = 55,
	ColumnStacked1003D = 56,
	Line3D = -4101,
	Pie3D = -4102,
	PieExploded3D = 70,
	Area = 1,
	AreaStacked = 76,
	AreaStacked100 = 77,
	BarClustered = 57,
	BarOfPie = 71,
	BarStacked = 58,
	BarStacked100 = 59,
	Bubble = 15,
	Bubble3DEffect = 87,
	ColumnClustered = 51,
	ColumnStacked = 52,
	ColumnStacked100 = 53,
	ConeBarClustered = 102,
	ConeBarStacked = 103,
	ConeBarStacked100 = 104,
	ConeCol = 105,
	ConeColClustered = 99,
	ConeColStacked = 100,
	ConeColStacked100 = 101,
	CylinderBarClustered = 95,
	CylinderBarStacked = 96,
	CylinderBarStacked100 = 97,
	CylinderCol = 98,
	CylinderColClustered = 92,
	CylinderColStacked = 93,
	CylinderColStacked100 = 94,
	Doughnut = -4120,
	DoughnutExploded = 80,
	Line = 4,
	LineMarkers = 65,
	LineMarkersStacked = 66,
	LineMarkersStacked100 = 67,
	LineStacked = 63,
	LineStacked100 = 64,
	Pie = 5,
	PieExploded = 69,
	PieOfPie = 68,
	PyramidBarClustered = 109,
	PyramidBarStacked = 110,
	PyramidBarStacked100 = 111,
	PyramidCol = 112,
	PyramidColClustered = 106,
	PyramidColStacked = 107,
	PyramidColStacked100 = 108,
	Radar = -4151,
	RadarFilled = 82,
	RadarMarkers = 81,
	StockHLC = 88,
	StockOHLC = 89,
	StockVHLC = 90,
	StockVOHLC = 91,
	Surface = 83,
	SurfaceTopView = 85,
	SurfaceTopViewWireframe = 86,
	SurfaceWireframe = 84,
	XYScatter = -4169,
	XYScatterLines = 74,
	XYScatterLinesNoMarkers = 75,
	XYScatterSmooth = 72,
	XYScatterSmoothNoMarkers = 73
}
/// <summary>
/// Bar or column
/// </summary>
public enum eDirection
{
	Column,
	Bar
}
/// <summary>
/// How the series are grouped
/// </summary>
public enum eGrouping
{
	Standard,
	Clustered,
	Stacked,
	PercentStacked
}
/// <summary>
/// Shape for bar charts
/// </summary>
public enum eShape
{
	Box,
	Cone,
	ConeToMax,
	Cylinder,
	Pyramid,
	PyramidToMax
}
/// <summary>
/// Smooth or lines markers
/// </summary>
public enum eScatterStyle
{
	LineMarker,
	SmoothMarker,
}
public enum eRadarStyle
{
	/// <summary>
	/// Specifies that the radar chart shall be filled and have lines but no markers.
	/// </summary>
	Filled,
	/// <summary>
	/// Specifies that the radar chart shall have lines and markers but no fill.
	/// </summary>
	Marker,
	/// <summary>
	/// Specifies that the radar chart shall have lines but no markers and no fill.
	/// </summary>
	Standard
}
/// <summary>
/// Bar or pie
/// </summary>
public enum ePieType
{
	Bar,
	Pie
}
/// <summary>
/// Position of the labels
/// </summary>
public enum eLabelPosition
{
	BestFit,
	Left,
	Right,
	Center,
	Top,
	Bottom,
	InBase,
	InEnd,
	OutEnd
}
/// <summary>
/// Axis label position
/// </summary>
public enum eTickLabelPosition
{
	High,
	Low,
	NextTo,
	None
}
/// <summary>
/// Markerstyle
/// </summary>
public enum eMarkerStyle
{
	Circle,
	Dash,
	Diamond,
	Dot,
	None,
	Picture,
	Plus,
	Square,
	Star,
	Triangle,
	X,
}
/// <summary>
/// The time unit of major and minor datetime axis values
/// </summary>
public enum eTimeUnit
{
	Years,
	Months,
	Days,
}
/// <summary>
/// The build in style of the chart.
/// </summary>
public enum eChartStyle
{
	None,
	Style1,
	Style2,
	Style3,
	Style4,
	Style5,
	Style6,
	Style7,
	Style8,
	Style9,
	Style10,
	Style11,
	Style12,
	Style13,
	Style14,
	Style15,
	Style16,
	Style17,
	Style18,
	Style19,
	Style20,
	Style21,
	Style22,
	Style23,
	Style24,
	Style25,
	Style26,
	Style27,
	Style28,
	Style29,
	Style30,
	Style31,
	Style32,
	Style33,
	Style34,
	Style35,
	Style36,
	Style37,
	Style38,
	Style39,
	Style40,
	Style41,
	Style42,
	Style43,
	Style44,
	Style45,
	Style46,
	Style47,
	Style48
}
/// <summary>
/// Type of Trendline for a chart
/// </summary>
public enum eTrendLine
{
	/// <summary>
	/// Specifies the trendline shall be an exponential curve in the form
	/// </summary>
	Exponential,
	/// <summary>
	/// Specifies the trendline shall be a logarithmic curve in the form , where log is the natural
	/// </summary>
	Linear,
	/// <summary>
	/// Specifies the trendline shall be a logarithmic curve in the form , where log is the natural
	/// </summary>
	Logarithmic,
	/// <summary>
	/// Specifies the trendline shall be a moving average of period Period
	/// </summary>
	MovingAvgerage,
	/// <summary>
	/// Specifies the trendline shall be a polynomial curve of order Order in the form 
	/// </summary>
	Polynomial,
	/// <summary>
	/// Specifies the trendline shall be a power curve in the form
	/// </summary>
	Power
}
/// <summary>
/// Specifies the possible ways to display blanks
/// </summary>
public enum eDisplayBlanksAs
{
	/// <summary>
	/// Blank values shall be left as a gap
	/// </summary>
	Gap,
	/// <summary>
	/// Blank values shall be spanned with a line (Line charts)
	/// </summary>
	Span,
	/// <summary>
	/// Blank values shall be treated as zero
	/// </summary>
	Zero
}
public enum eSizeRepresents
{
	/// <summary>
	/// Specifies the area of the bubbles shall be proportional to the bubble size value.
	/// </summary>
	Area,
	/// <summary>
	/// Specifies the radius of the bubbles shall be proportional to the bubble size value.
	/// </summary>
	Width
}
#endregion
/// <summary>
/// Base class for Chart object.
/// </summary>
public class ExcelChart : ExcelDrawing
{
	const string rootPath = "c:chartSpace/c:chart/c:plotArea";
	//string _chartPath;
	protected internal ExcelChartSeries _chartSeries;
	internal ExcelChartAxis[] _axis;
	protected XmlHelper _chartXmlHelper;
	#region "Constructors"
	internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot) :
		base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
	{
		ChartType = type;
		CreateNewChart(drawings, type, null);

		Init(drawings, _chartNode);

		_chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode, isPivot);

		SetTypeProperties();
		LoadAxis();
	}
	internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
		base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
	{
		ChartType = type;
		CreateNewChart(drawings, type, topChart);

		Init(drawings, _chartNode);

		_chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode, PivotTableSource != null);
		if (PivotTableSource != null) SetPivotSource(PivotTableSource);

		SetTypeProperties();
		if (topChart == null)
			LoadAxis();
		else
		{
			_axis = topChart.Axis;
			if (_axis.Length > 0)
			{
				XAxis = _axis[0];
				YAxis = _axis[1];
			}
		}
	}
	internal ExcelChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
		base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
	{
		UriChart = uriChart;
		Part = part;
		ChartXml = chartXml;
		_chartNode = chartNode;
		InitChartLoad(drawings, chartNode);
		ChartType = GetChartType(chartNode.LocalName);
	}
	internal ExcelChart(ExcelChart topChart, XmlNode chartNode) :
		base(topChart._drawings, topChart.TopNode, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
	{
		UriChart = topChart.UriChart;
		Part = topChart.Part;
		ChartXml = topChart.ChartXml;
		_plotArea = topChart.PlotArea;
		_chartNode = chartNode;

		InitChartLoad(topChart._drawings, chartNode);
	}
	private void InitChartLoad(ExcelDrawings drawings, XmlNode chartNode)
	{
		//SetChartType();
		var isPivot = false;
		Init(drawings, chartNode);
		_chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode, isPivot /*ChartXml.SelectSingleNode(_chartPath, drawings.NameSpaceManager)*/);
		LoadAxis();
	}

	private void Init(ExcelDrawings drawings, XmlNode chartNode)
	{
		//_chartXmlHelper = new XmlHelper(drawings.NameSpaceManager, chartNode);
		_chartXmlHelper = XmlHelperFactory.Create(drawings.NameSpaceManager, chartNode);
		_chartXmlHelper.SchemaNodeOrder = ["ofPieType", "title", "pivotFmt", "autoTitleDeleted", "view3D", "floor", "sideWall", "backWall", "plotArea", "wireframe", "barDir", "grouping", "scatterStyle", "radarStyle", "varyColors", "ser", "dLbls", "bubbleScale", "showNegBubbles", "dropLines", "upDownBars", "marker", "smooth", "shape", "legend", "plotVisOnly", "dispBlanksAs", "gapWidth", "showDLblsOverMax", "overlap", "bandFmts", "axId", "spPr", "printSettings"];
		WorkSheet = drawings.Worksheet;
	}
	#endregion
	#region "Private functions"
	private void SetTypeProperties()
	{
		/******* Grouping *******/
		if (IsTypeClustered())
		{
			Grouping = eGrouping.Clustered;
		}
		else if (
			IsTypeStacked())
		{
			Grouping = eGrouping.Stacked;
		}
		else if (
		IsTypePercentStacked())
		{
			Grouping = eGrouping.PercentStacked;
		}

		/***** 3D Perspective *****/
		if (IsType3D())
		{
			View3D.RotY = 20;
			View3D.Perspective = 30;    //Default to 30
			if (IsTypePieDoughnut())
			{
				View3D.RotX = 30;
			}
			else
			{
				View3D.RotX = 15;
			}
		}
	}
	private void CreateNewChart(ExcelDrawings drawings, eChartType type, ExcelChart topChart)
	{
		if (topChart == null)
		{
			var graphFrame = TopNode.OwnerDocument.CreateElement("graphicFrame", ExcelPackage.schemaSheetDrawings);
			graphFrame.SetAttribute("macro", "");
			TopNode.AppendChild(graphFrame);
			graphFrame.InnerXml = string.Format("<xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"Chart 1\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\" /> <a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" />   </a:graphicData>  </a:graphic>", _id);
			TopNode.AppendChild(TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

			var package = drawings.Worksheet._package.Package;
			UriChart = GetNewUri(package, "/xl/charts/chart{0}.xml");

			ChartXml = new XmlDocument
			{
				PreserveWhitespace = ExcelPackage.preserveWhitespace
			};
			LoadXmlSafe(ChartXml, ChartStartXml(type), Encoding.UTF8);

			// save it to the package
			Part = package.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _drawings._package.Compression);

			StreamWriter streamChart = new(Part.GetStream(FileMode.Create, FileAccess.Write));
			ChartXml.Save(streamChart);
#if !Core
			streamChart.Close();
#endif
			package.Flush();

			var chartRelation = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
			graphFrame.SelectSingleNode("a:graphic/a:graphicData/c:chart", NameSpaceManager).Attributes["r:id"].Value = chartRelation.Id;
			package.Flush();
			_chartNode = ChartXml.SelectSingleNode(string.Format("c:chartSpace/c:chart/c:plotArea/{0}", GetChartNodeText()), NameSpaceManager);
		}
		else
		{
			ChartXml = topChart.ChartXml;
			Part = topChart.Part;
			_plotArea = topChart.PlotArea;
			UriChart = topChart.UriChart;
			_axis = topChart._axis;

			var preNode = _plotArea.ChartTypes[^1].ChartNode;
			_chartNode = ((XmlDocument)ChartXml).CreateElement(GetChartNodeText(), ExcelPackage.schemaChart);
			preNode.ParentNode.InsertAfter(_chartNode, preNode);
			if (topChart.Axis.Length == 0)
			{
				AddAxis();
			}

			var serieXML = GetChartSerieStartXml(type, int.Parse(topChart.Axis[0].Id), int.Parse(topChart.Axis[1].Id), topChart.Axis.Length > 2 ? int.Parse(topChart.Axis[2].Id) : -1);
			_chartNode.InnerXml = serieXML;
		}

		GetPositionSize();
	}
	private void LoadAxis()
	{
		var nl = _chartNode.SelectNodes("c:axId", NameSpaceManager);
		List<ExcelChartAxis> l = [];
		foreach (XmlNode node in nl)
		{
			var id = node.Attributes["val"].Value;
			var axNode = ChartXml.SelectNodes(rootPath + string.Format("/*/c:axId[@val=\"{0}\"]", id), NameSpaceManager);
			if (axNode != null && axNode.Count > 1)
			{
				foreach (XmlNode axn in axNode)
				{
					if (axn.ParentNode.LocalName.EndsWith("Ax"))
					{
						var axisNode = axNode[1].ParentNode;
						ExcelChartAxis ax = new(NameSpaceManager, axisNode);
						l.Add(ax);
					}
				}
			}
		}

		_axis = [.. l];

		if (_axis.Length > 0) XAxis = _axis[0];
		if (_axis.Length > 1) YAxis = _axis[1];
	}
	//private void SetChartType()
	//{
	//    ChartType = 0;
	//    //_plotArea = new ExcelChartPlotArea(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager));
	//    int pos=0;
	//    foreach (XmlElement n in ChartXml.SelectSingleNode(rootPath, _drawings.NameSpaceManager).ChildNodes)
	//    {
	//        if (pos == 0)
	//        {
	//            ChartType = GetChartType(n.Name);
	//            if (ChartType != 0)
	//            {
	//                //_chartPath = rootPath + "/" + n.Name;
	//                PlotArea.ChartTypes.Add(this);
	//            }
	//        }
	//        else
	//        {
	//            var chartSerieType = GetChart(_drawings, TopNode/*, n*/);
	//            chartSerieType = GetChart(n, _drawings, TopNode, UriChart, Part, ChartXml, null, isPivot);
	//            PlotArea.ChartTypes.Add(chartSerieType);
	//            //var chartType = GetChartType(n.Name);
	//        }
	//        if (ChartType != 0)
	//        {
	//            pos++;
	//        }
	//    }
	//}
	internal virtual eChartType GetChartType(string name)
	{

		switch (name)
		{
			case "area3DChart":
				if (Grouping == eGrouping.Stacked)
				{
					return eChartType.AreaStacked3D;
				}
				else
				{
					return Grouping == eGrouping.PercentStacked ? eChartType.AreaStacked1003D : eChartType.Area3D;
				}
			case "areaChart":
				if (Grouping == eGrouping.Stacked)
				{
					return eChartType.AreaStacked;
				}
				else
				{
					return Grouping == eGrouping.PercentStacked ? eChartType.AreaStacked100 : eChartType.Area;
				}
			case "doughnutChart":
				return eChartType.Doughnut;
			case "pie3DChart":
				return eChartType.Pie3D;
			case "pieChart":
				return eChartType.Pie;
			case "radarChart":
				return eChartType.Radar;
			case "scatterChart":
				return eChartType.XYScatter;
			case "surface3DChart":
			case "surfaceChart":
				return eChartType.Surface;
			case "stockChart":
				return eChartType.StockHLC;
			default:
				return 0;
		}
	}
	#region "Xml init Functions"
	private string ChartStartXml(eChartType type)
	{
		StringBuilder xml = new();
		var axID = 1;
		var xAxID = 2;
		var serAxID = IsTypeSurface() ? 3 : -1;

		xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
		xml.AppendFormat("<c:chartSpace xmlns:c=\"{0}\" xmlns:a=\"{1}\" xmlns:r=\"{2}\">", ExcelPackage.schemaChart, ExcelPackage.schemaDrawings, ExcelPackage.schemaRelationships);
		xml.Append("<c:chart>");
		xml.AppendFormat("{0}{1}<c:plotArea><c:layout/>", AddPerspectiveXml(type), AddSurfaceXml(type));

		var chartNodeText = GetChartNodeText();
		xml.AppendFormat("<{0}>", chartNodeText);
		xml.Append(GetChartSerieStartXml(type, axID, xAxID, serAxID));
		xml.AppendFormat("</{0}>", chartNodeText);

		//Axis
		if (!IsTypePieDoughnut())
		{
			if (IsTypeScatterBubble())
			{
				xml.AppendFormat("<c:valAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:valAx>", axID, xAxID);
			}
			else
			{
				xml.AppendFormat("<c:catAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>", axID, xAxID);
			}

			xml.AppendFormat("<c:valAx><c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>", axID, xAxID);
			if (serAxID == 3) //Sureface Chart
			{
				xml.AppendFormat("<c:serAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:serAx>", serAxID, xAxID);
			}
		}

		xml.AppendFormat("</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"0\" /></c:legend><c:plotVisOnly val=\"1\"/></c:chart>", axID, xAxID);

		xml.Append("<c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings></c:chartSpace>");
		return xml.ToString();
	}

	private string GetChartSerieStartXml(eChartType type, int axID, int xAxID, int serAxID)
	{
		StringBuilder xml = new();

		xml.Append(AddScatterType(type));
		xml.Append(AddRadarType(type));
		xml.Append(AddBarDir(type));
		xml.Append(AddGrouping());
		xml.Append(AddVaryColors());
		xml.Append(AddHasMarker(type));
		xml.Append(AddShape(type));
		xml.Append(AddFirstSliceAng(type));
		xml.Append(AddHoleSize(type));
		if (ChartType is eChartType.BarStacked100 or
			eChartType.BarStacked or
			eChartType.ColumnStacked or
			eChartType.ColumnStacked100)
		{
			xml.Append("<c:overlap val=\"100\"/>");
		}

		if (IsTypeSurface())
		{
			xml.Append("<c:bandFmts/>");
		}

		xml.Append(AddAxisId(axID, xAxID, serAxID));

		return xml.ToString();
	}
	private string AddAxisId(int axID, int xAxID, int serAxID)
	{
		if (!IsTypePieDoughnut())
		{
			return IsTypeSurface()
				? string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/><c:axId val=\"{2}\"/>", axID, xAxID, serAxID)
				: string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/>", axID, xAxID);
		}
		else
		{
			return "";
		}
	}
	private string AddAxType() => ChartType switch
	{
		eChartType.XYScatter or eChartType.XYScatterLines or eChartType.XYScatterLinesNoMarkers or eChartType.XYScatterSmooth or eChartType.XYScatterSmoothNoMarkers or eChartType.Bubble or eChartType.Bubble3DEffect => "valAx",
		_ => "catAx",
	};
	private string AddScatterType(eChartType type) => type is eChartType.XYScatter or
			eChartType.XYScatterLines or
			eChartType.XYScatterLinesNoMarkers or
			eChartType.XYScatterSmooth or
			eChartType.XYScatterSmoothNoMarkers
			? "<c:scatterStyle val=\"\" />"
			: "";
	private string AddRadarType(eChartType type) => type is eChartType.Radar or
			eChartType.RadarFilled or
			eChartType.RadarMarkers
			? "<c:radarStyle val=\"\" />"
			: "";
	private string AddGrouping() =>
		//IsTypeClustered() || IsTypePercentStacked() || IsTypeStacked() || 
		IsTypeShape() || IsTypeLine() ? "<c:grouping val=\"standard\"/>" : "";
	private string AddHoleSize(eChartType type) => type is eChartType.Doughnut or
			eChartType.DoughnutExploded
			? "<c:holeSize val=\"50\" />"
			: "";
	private string AddFirstSliceAng(eChartType type) => type is eChartType.Doughnut or
			eChartType.DoughnutExploded
			? "<c:firstSliceAng val=\"0\" />"
			: "";
	private string AddVaryColors() => IsTypePieDoughnut() ? "<c:varyColors val=\"1\" />" : "<c:varyColors val=\"0\" />";
	private string AddHasMarker(eChartType type) => type is eChartType.LineMarkers or
			eChartType.LineMarkersStacked or
			eChartType.LineMarkersStacked100
			? "<c:marker val=\"1\"/>"
			: "";
	private string AddShape(eChartType type) => IsTypeShape() ? "<c:shape val=\"box\" />" : "";
	private string AddBarDir(eChartType type) => IsTypeShape() ? "<c:barDir val=\"col\" />" : "";
	private string AddPerspectiveXml(eChartType type) =>
		//Add for 3D sharts
		IsType3D() ? "<c:view3D><c:perspective val=\"30\" /></c:view3D>" : "";
	private string AddSurfaceXml(eChartType type) => IsTypeSurface() ? AddSurfacePart("floor") + AddSurfacePart("sideWall") + AddSurfacePart("backWall") : "";

	private string AddSurfacePart(string name) => string.Format("<c:{0}><c:thickness val=\"0\"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/></c:spPr></c:{0}>", name);
	#endregion
	#endregion
	#region "Chart type functions
	internal static bool IsType3D(eChartType chartType) => chartType is eChartType.Area3D or
						eChartType.AreaStacked3D or
						eChartType.AreaStacked1003D or
						eChartType.BarClustered3D or
						eChartType.BarStacked3D or
						eChartType.BarStacked1003D or
						eChartType.Column3D or
						eChartType.ColumnClustered3D or
						eChartType.ColumnStacked3D or
						eChartType.ColumnStacked1003D or
						eChartType.Line3D or
						eChartType.Pie3D or
						eChartType.PieExploded3D or
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
						eChartType.PyramidColStacked100 or
						eChartType.Surface or
						eChartType.SurfaceTopView or
						eChartType.SurfaceTopViewWireframe or
						eChartType.SurfaceWireframe;
	internal protected bool IsType3D() => IsType3D(ChartType);
	protected bool IsTypeLine() => ChartType is eChartType.Line or
				eChartType.LineMarkers or
				eChartType.LineMarkersStacked100 or
				eChartType.LineStacked or
				eChartType.LineStacked100 or
				eChartType.Line3D;
	protected bool IsTypeScatterBubble() => ChartType is eChartType.XYScatter or
				eChartType.XYScatterLines or
				eChartType.XYScatterLinesNoMarkers or
				eChartType.XYScatterSmooth or
				eChartType.XYScatterSmoothNoMarkers or
				eChartType.Bubble or
				eChartType.Bubble3DEffect;
	protected bool IsTypeSurface() => ChartType is eChartType.Surface or
			   eChartType.SurfaceTopView or
			   eChartType.SurfaceTopViewWireframe or
			   eChartType.SurfaceWireframe;
	protected bool IsTypeShape() => ChartType is eChartType.BarClustered3D or
				eChartType.BarStacked3D or
				eChartType.BarStacked1003D or
				eChartType.BarClustered3D or
				eChartType.BarStacked3D or
				eChartType.BarStacked1003D or
				eChartType.Column3D or
				eChartType.ColumnClustered3D or
				eChartType.ColumnStacked3D or
				eChartType.ColumnStacked1003D or
				//ChartType == eChartType.3DPie ||
				//ChartType == eChartType.3DPieExploded ||
				//ChartType == eChartType.Bubble3DEffect ||
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
				eChartType.PyramidColStacked100; //||//ChartType == eChartType.Doughnut ||//ChartType == eChartType.DoughnutExploded;
	protected internal bool IsTypePercentStacked() => ChartType is eChartType.AreaStacked100 or
					   eChartType.BarStacked100 or
					   eChartType.BarStacked1003D or
					   eChartType.ColumnStacked100 or
					   eChartType.ColumnStacked1003D or
					   eChartType.ConeBarStacked100 or
					   eChartType.ConeColStacked100 or
					   eChartType.CylinderBarStacked100 or
					   eChartType.CylinderColStacked or
					   eChartType.LineMarkersStacked100 or
					   eChartType.LineStacked100 or
					   eChartType.PyramidBarStacked100 or
					   eChartType.PyramidColStacked100;
	protected internal bool IsTypeStacked() => ChartType is eChartType.AreaStacked or
					   eChartType.AreaStacked3D or
					   eChartType.BarStacked or
					   eChartType.BarStacked3D or
					   eChartType.ColumnStacked3D or
					   eChartType.ColumnStacked or
					   eChartType.ConeBarStacked or
					   eChartType.ConeColStacked or
					   eChartType.CylinderBarStacked or
					   eChartType.CylinderColStacked or
					   eChartType.LineMarkersStacked or
					   eChartType.LineStacked or
					   eChartType.PyramidBarStacked or
					   eChartType.PyramidColStacked;
	protected bool IsTypeClustered() => ChartType is eChartType.BarClustered or
					   eChartType.BarClustered3D or
					   eChartType.ColumnClustered3D or
					   eChartType.ColumnClustered or
					   eChartType.ConeBarClustered or
					   eChartType.ConeColClustered or
					   eChartType.CylinderBarClustered or
					   eChartType.CylinderColClustered or
					   eChartType.PyramidBarClustered or
					   eChartType.PyramidColClustered;
	protected internal bool IsTypePieDoughnut() => ChartType is eChartType.Pie or
					   eChartType.PieExploded or
					   eChartType.PieOfPie or
					   eChartType.Pie3D or
					   eChartType.PieExploded3D or
					   eChartType.BarOfPie or
					   eChartType.Doughnut or
					   eChartType.DoughnutExploded;
	#endregion
	/// <summary>
	/// Get the name of the chart node
	/// </summary>
	/// <returns>The name</returns>
	protected string GetChartNodeText() => ChartType switch
	{
		eChartType.Area3D or eChartType.AreaStacked3D or eChartType.AreaStacked1003D => "c:area3DChart",
		eChartType.Area or eChartType.AreaStacked or eChartType.AreaStacked100 => "c:areaChart",
		eChartType.BarClustered or eChartType.BarStacked or eChartType.BarStacked100 or eChartType.ColumnClustered or eChartType.ColumnStacked or eChartType.ColumnStacked100 => "c:barChart",
		eChartType.BarClustered3D or eChartType.BarStacked3D or eChartType.BarStacked1003D or eChartType.ColumnClustered3D or eChartType.ColumnStacked3D or eChartType.ColumnStacked1003D or eChartType.ConeBarClustered or eChartType.ConeBarStacked or eChartType.ConeBarStacked100 or eChartType.ConeCol or eChartType.ConeColClustered or eChartType.ConeColStacked or eChartType.ConeColStacked100 or eChartType.CylinderBarClustered or eChartType.CylinderBarStacked or eChartType.CylinderBarStacked100 or eChartType.CylinderCol or eChartType.CylinderColClustered or eChartType.CylinderColStacked or eChartType.CylinderColStacked100 or eChartType.PyramidBarClustered or eChartType.PyramidBarStacked or eChartType.PyramidBarStacked100 or eChartType.PyramidCol or eChartType.PyramidColClustered or eChartType.PyramidColStacked or eChartType.PyramidColStacked100 => "c:bar3DChart",
		eChartType.Bubble or eChartType.Bubble3DEffect => "c:bubbleChart",
		eChartType.Doughnut or eChartType.DoughnutExploded => "c:doughnutChart",
		eChartType.Line or eChartType.LineMarkers or eChartType.LineMarkersStacked or eChartType.LineMarkersStacked100 or eChartType.LineStacked or eChartType.LineStacked100 => "c:lineChart",
		eChartType.Line3D => "c:line3DChart",
		eChartType.Pie or eChartType.PieExploded => "c:pieChart",
		eChartType.BarOfPie or eChartType.PieOfPie => "c:ofPieChart",
		eChartType.Pie3D or eChartType.PieExploded3D => "c:pie3DChart",
		eChartType.Radar or eChartType.RadarFilled or eChartType.RadarMarkers => "c:radarChart",
		eChartType.XYScatter or eChartType.XYScatterLines or eChartType.XYScatterLinesNoMarkers or eChartType.XYScatterSmooth or eChartType.XYScatterSmoothNoMarkers => "c:scatterChart",
		eChartType.Surface or eChartType.SurfaceWireframe => "c:surface3DChart",
		eChartType.SurfaceTopView or eChartType.SurfaceTopViewWireframe => "c:surfaceChart",
		eChartType.StockHLC => "c:stockChart",
		_ => throw (new NotImplementedException("Chart type not implemented")),
	};
	/// <summary>
	/// Add a secondary axis
	/// </summary>
	internal void AddAxis()
	{
		var catAx = ChartXml.CreateElement(string.Format("c:{0}", AddAxType()), ExcelPackage.schemaChart);
		int axID;
		if (_axis.Length == 0)
		{
			_plotArea.TopNode.AppendChild(catAx);
			axID = 1;
		}
		else
		{
			_axis[0].TopNode.ParentNode.InsertAfter(catAx, _axis[^1].TopNode);
			axID = int.Parse(_axis[0].Id) < int.Parse(_axis[1].Id) ? int.Parse(_axis[1].Id) + 1 : int.Parse(_axis[0].Id) + 1;
		}


		var valAx = ChartXml.CreateElement("c:valAx", ExcelPackage.schemaChart);
		catAx.ParentNode.InsertAfter(valAx, catAx);

		if (_axis.Length == 0)
		{
			catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/>", axID, axID + 1);
			valAx.InnerXml = string.Format("<c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/>", axID, axID + 1);
		}
		else
		{
			catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"1\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"none\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/>", axID, axID + 1);
			valAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"r\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/>", axID + 1, axID);
		}

		if (_axis.Length == 0)
		{
			_axis = new ExcelChartAxis[2];
		}
		else
		{
			var newAxis = new ExcelChartAxis[_axis.Length + 2];
			Array.Copy(_axis, newAxis, _axis.Length);
			_axis = newAxis;
		}

		_axis[^2] = new ExcelChartAxis(NameSpaceManager, catAx);
		_axis[^1] = new ExcelChartAxis(NameSpaceManager, valAx);
		foreach (var chart in _plotArea.ChartTypes)
		{
			chart._axis = _axis;
		}
	}
	internal void RemoveSecondaryAxis() => throw (new NotImplementedException("Not yet implemented"));
	#region "Properties"
	/// <summary>
	/// Reference to the worksheet
	/// </summary>
	public ExcelWorksheet WorkSheet { get; internal set; }
	/// <summary>
	/// The chart xml document
	/// </summary>
	public XmlDocument ChartXml { get; internal set; }
	/// <summary>
	/// Type of chart
	/// </summary>
	public eChartType ChartType { get; internal set; }
	internal protected XmlNode _chartNode = null;
	internal XmlNode ChartNode => _chartNode;
	/// <summary>
	/// Titel of the chart
	/// </summary>
	public ExcelChartTitle Title
	{
		get
		{
			_title ??= new ExcelChartTitle(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart", NameSpaceManager));

			return _title;
		}
	}
	/// <summary>
	/// Chart series
	/// </summary>
	public virtual ExcelChartSeries Series => _chartSeries;
	/// <summary>
	/// An array containg all axis of all Charttypes
	/// </summary>
	public ExcelChartAxis[] Axis => _axis;
	/// <summary>
	/// The XAxis
	/// </summary>
	public ExcelChartAxis XAxis
	{
		get;
		private set;
	}
	/// <summary>
	/// The YAxis
	/// </summary>
	public ExcelChartAxis YAxis
	{
		get;
		private set;
	}
	bool _secondaryAxis = false;
	/// <summary>
	/// If true the charttype will use the secondary axis.
	/// The chart must contain a least one other charttype that uses the primary axis.
	/// </summary>
	public bool UseSecondaryAxis
	{
		get
		{
			return _secondaryAxis;
		}
		set
		{
			if (_secondaryAxis != value)
			{
				if (value)
				{
					if (IsTypePieDoughnut())
					{
						throw (new Exception("Pie charts do not support axis"));
					}
					else if (HasPrimaryAxis() == false)
					{
						throw (new Exception("Can't set to secondary axis when no serie uses the primary axis"));
					}

					if (Axis.Length == 2)
					{
						AddAxis();
					}

					var nl = ChartNode.SelectNodes("c:axId", NameSpaceManager);
					nl[0].Attributes["val"].Value = Axis[2].Id;
					nl[1].Attributes["val"].Value = Axis[3].Id;
					XAxis = Axis[2];
					YAxis = Axis[3];
				}
				else
				{
					var nl = ChartNode.SelectNodes("c:axId", NameSpaceManager);
					nl[0].Attributes["val"].Value = Axis[0].Id;
					nl[1].Attributes["val"].Value = Axis[1].Id;
					XAxis = Axis[0];
					YAxis = Axis[1];
				}

				_secondaryAxis = value;
			}
		}
	}
	/// <summary>
	/// The build-in chart styles. 
	/// </summary>
	public eChartStyle Style
	{
		get
		{
			var node = ChartXml.SelectSingleNode("c:chartSpace/c:style/@val", NameSpaceManager);
			if (node == null)
			{
				return eChartStyle.None;
			}
			else
			{
				return int.TryParse(node.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var v) ? (eChartStyle)v : eChartStyle.None;
			}

		}
		set
		{
			if (value == eChartStyle.None)
			{
				var element = ChartXml.SelectSingleNode("c:chartSpace/c:style", NameSpaceManager) as XmlElement;
				element?.ParentNode.RemoveChild(element);
			}
			else
			{
				var element = ChartXml.CreateElement("c:style", ExcelPackage.schemaChart);
				element.SetAttribute("val", ((int)value).ToString());
				var parent = ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager) as XmlElement;
				parent.InsertBefore(element, parent.SelectSingleNode("c:chart", NameSpaceManager));
			}
		}
	}
	const string _roundedCornersPath = "../../../c:roundedCorners/@val";
	/// <summary>
	/// Border rounded corners
	/// </summary>
	public bool RoundedCorners
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool(_roundedCornersPath);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool(_roundedCornersPath, value);
		}
	}
	const string _plotVisibleOnlyPath = "../../c:plotVisOnly/@val";
	/// <summary>
	/// Show data in hidden rows and columns
	/// </summary>
	public bool ShowHiddenData
	{
		get
		{
			//!!Inverted value!!
			return !_chartXmlHelper.GetXmlNodeBool(_plotVisibleOnlyPath);
		}
		set
		{
			//!!Inverted value!!
			_chartXmlHelper.SetXmlNodeBool(_plotVisibleOnlyPath, !value);
		}
	}
	const string _displayBlanksAsPath = "../../c:dispBlanksAs/@val";
	/// <summary>
	/// Specifies the possible ways to display blanks
	/// </summary>
	public eDisplayBlanksAs DisplayBlanksAs
	{
		get
		{
			var v = _chartXmlHelper.GetXmlNodeString(_displayBlanksAsPath);
			if (string.IsNullOrEmpty(v))
			{
				return eDisplayBlanksAs.Zero; //Issue 14715 Changed in Office 2010-?
			}
			else
			{
				return (eDisplayBlanksAs)Enum.Parse(typeof(eDisplayBlanksAs), v, true);
			}
		}
		set
		{
			_chartSeries.SetXmlNodeString(_displayBlanksAsPath, value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}
	const string _showDLblsOverMax = "../../c:showDLblsOverMax/@val";
	/// <summary>
	/// Specifies data labels over the maximum of the chart shall be shown
	/// </summary>
	public bool ShowDataLabelsOverMaximum
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool(_showDLblsOverMax, true);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool(_showDLblsOverMax, value, true);
		}
	}
	private bool HasPrimaryAxis()
	{
		if (_plotArea.ChartTypes.Count == 1)
		{
			return false;
		}

		foreach (var chart in _plotArea.ChartTypes)
		{
			if (chart != this)
			{
				if (chart.UseSecondaryAxis == false && chart.IsTypePieDoughnut() == false)
				{
					return true;
				}
			}
		}

		return false;
	}
	///// <summary>
	///// Sets position of the axis of a chart-serie
	///// </summary>
	///// <param name="XAxis">Left or Right</param>
	///// <param name="YAxis">Top or Bottom</param>
	//internal void SetAxis(eXAxisPosition XAxis, eYAxisPosition YAxis)
	//{
	//    bool xAxisExists = false, yAxisExists = false;
	//    foreach (var axis in _axis)
	//    {
	//        if (axis.AxisPosition == (eAxisPosition)XAxis)
	//        {
	//            //Found
	//            xAxisExists=true;
	//            if (axis != this.XAxis)
	//            {
	//                CheckRemoveAxis(this.XAxis);
	//                this.XAxis = axis;
	//            }
	//        }
	//        else if (axis.AxisPosition == (eAxisPosition)YAxis)
	//        {
	//            yAxisExists = true;
	//            if (axis != this.YAxis)
	//            {
	//                CheckRemoveAxis(this.YAxis);
	//                this.YAxis = axis;
	//            }
	//        }
	//    }

	//    if (!xAxisExists)
	//    {
	//        if (ExistsAxis(this.XAxis))
	//        {
	//            AddAxis((eAxisPosition)XAxis);
	//            this.XAxis = Axis[Axis.Length - 1];
	//        }
	//        else
	//        {
	//            this.XAxis.AxisPosition = (eAxisPosition)XAxis;
	//        }
	//    }
	//    if (!yAxisExists)
	//    {
	//        if (ExistsAxis(this.XAxis))
	//        {
	//            AddAxis((eAxisPosition)YAxis);
	//            this.YAxis = Axis[Axis.Length - 1];
	//        }
	//        else
	//        {
	//            this.YAxis.AxisPosition = (eAxisPosition)YAxis;
	//        }
	//    }
	//}

	/// <summary>
	/// Remove all axis that are not used any more
	/// </summary>
	/// <param name="excelChartAxis"></param>
	private void CheckRemoveAxis(ExcelChartAxis excelChartAxis)
	{
		if (ExistsAxis(excelChartAxis))
		{
			//Remove the axis
			var newAxis = new ExcelChartAxis[Axis.Length - 1];
			var pos = 0;
			foreach (var ax in Axis)
			{
				if (ax != excelChartAxis)
				{
					newAxis[pos] = ax;
				}
			}

			//Update all charttypes.
			foreach (var chartType in _plotArea.ChartTypes)
			{
				chartType._axis = newAxis;
			}
		}
	}

	private bool ExistsAxis(ExcelChartAxis excelChartAxis)
	{
		foreach (var chartType in _plotArea.ChartTypes)
		{
			if (chartType != this)
			{
				if (chartType.XAxis.AxisPosition == excelChartAxis.AxisPosition ||
				   chartType.YAxis.AxisPosition == excelChartAxis.AxisPosition)
				{
					//The axis exists
					return true;
				}
			}
		}

		return false;
	}
	ExcelChartPlotArea _plotArea = null;
	/// <summary>
	/// Plotarea
	/// </summary>
	public ExcelChartPlotArea PlotArea
	{
		get
		{
			_plotArea ??= new ExcelChartPlotArea(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager), this);

			return _plotArea;
		}
	}
	ExcelChartLegend _legend = null;
	/// <summary>
	/// Legend
	/// </summary>
	public ExcelChartLegend Legend
	{
		get
		{
			_legend ??= new ExcelChartLegend(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager), this);

			return _legend;
		}

	}
	ExcelDrawingBorder _border = null;
	/// <summary>
	/// Border
	/// </summary>
	public ExcelDrawingBorder Border
	{
		get
		{
			_border ??= new ExcelDrawingBorder(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr/a:ln");

			return _border;
		}
	}
	ExcelDrawingFill _fill = null;
	/// <summary>
	/// Fill
	/// </summary>
	public ExcelDrawingFill Fill
	{
		get
		{
			_fill ??= new ExcelDrawingFill(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr");

			return _fill;
		}
	}
	/// <summary>
	/// 3D-settings
	/// </summary>
	public ExcelView3D View3D => IsType3D()
				? new ExcelView3D(NameSpaceManager, ChartXml.SelectSingleNode("//c:view3D", NameSpaceManager))
				: throw (new Exception("Charttype does not support 3D"));
	//string _groupingPath = "c:chartSpace/c:chart/c:plotArea/{0}/c:grouping/@val";
	string _groupingPath = "c:grouping/@val";
	public eGrouping Grouping
	{
		get
		{
			return GetGroupingEnum(_chartXmlHelper.GetXmlNodeString(_groupingPath));
		}
		internal set
		{
			_chartXmlHelper.SetXmlNodeString(_groupingPath, GetGroupingText(value));
		}
	}
	//string _varyColorsPath = "c:chartSpace/c:chart/c:plotArea/{0}/c:varyColors/@val";
	string _varyColorsPath = "c:varyColors/@val";
	/// <summary>
	/// If the chart has only one serie this varies the colors for each point.
	/// </summary>
	public bool VaryColors
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool(_varyColorsPath);
		}
		set
		{
			if (value)
			{
				_chartXmlHelper.SetXmlNodeString(_varyColorsPath, "1");
			}
			else
			{
				_chartXmlHelper.SetXmlNodeString(_varyColorsPath, "0");
			}
		}
	}
	internal Packaging.ZipPackagePart Part { get; set; }
	/// <summary>
	/// Package internal URI
	/// </summary>
	internal Uri UriChart { get; set; }
	internal new string Id => "";
	ExcelChartTitle _title = null;
	#endregion
	#region "Grouping Enum Translation"
	private string GetGroupingText(eGrouping grouping) => grouping switch
	{
		eGrouping.Clustered => "clustered",
		eGrouping.Stacked => "stacked",
		eGrouping.PercentStacked => "percentStacked",
		_ => "standard",
	};
	private eGrouping GetGroupingEnum(string grouping) => grouping switch
	{
		"stacked" => eGrouping.Stacked,
		"percentStacked" => eGrouping.PercentStacked,
		//"clustered":               
		_ => eGrouping.Clustered,
	};
	#endregion
	internal static ExcelChart GetChart(ExcelDrawings drawings, XmlNode node/*, XmlNode chartTypeNode*/)
	{
		var chartNode = node.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart", drawings.NameSpaceManager);
		if (chartNode != null)
		{
			var drawingRelation = drawings.Part.GetRelationship(chartNode.Attributes["r:id"].Value);
			var uriChart = UriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

			var part = drawings.Part.Package.GetPart(uriChart);
			var chartXml = new XmlDocument();
			LoadXmlSafe(chartXml, part.GetStream());

			ExcelChart topChart = null;
			foreach (XmlElement n in chartXml.SelectSingleNode(rootPath, drawings.NameSpaceManager).ChildNodes)
			{
				if (topChart == null)
				{
					topChart = GetChart(n, drawings, node, uriChart, part, chartXml, null);
					topChart?.PlotArea.ChartTypes.Add(topChart);
				}
				else
				{
					var subChart = GetChart(n, null, null, null, null, null, topChart);
					if (subChart != null)
					{
						topChart.PlotArea.ChartTypes.Add(subChart);
					}
				}
			}

			return topChart;
		}
		else
		{
			return null;
		}
	}
	internal static ExcelChart GetChart(XmlElement chartNode, ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, ExcelChart topChart)
	{
		switch (chartNode.LocalName)
		{
			case "area3DChart":
			case "areaChart":
			case "stockChart":
				return topChart == null ? new ExcelChart(drawings, node, uriChart, part, chartXml, chartNode) : new ExcelChart(topChart, chartNode);
			case "surface3DChart":
			case "surfaceChart":
				return topChart == null
					? new ExcelSurfaceChart(drawings, node, uriChart, part, chartXml, chartNode)
					: (ExcelChart)new ExcelSurfaceChart(topChart, chartNode);
			case "radarChart":
				return topChart == null
					? new ExcelRadarChart(drawings, node, uriChart, part, chartXml, chartNode)
					: (ExcelChart)new ExcelRadarChart(topChart, chartNode);
			case "bubbleChart":
				return topChart == null
					? new ExcelBubbleChart(drawings, node, uriChart, part, chartXml, chartNode)
					: (ExcelChart)new ExcelBubbleChart(topChart, chartNode);
			case "barChart":
			case "bar3DChart":
				return topChart == null
					? new ExcelBarChart(drawings, node, uriChart, part, chartXml, chartNode)
					: (ExcelChart)new ExcelBarChart(topChart, chartNode);
			case "doughnutChart":
				return topChart == null
					? new ExcelDoughnutChart(drawings, node, uriChart, part, chartXml, chartNode)
					: (ExcelChart)new ExcelDoughnutChart(topChart, chartNode);
			case "pie3DChart":
			case "pieChart":
				return topChart == null
					? new ExcelPieChart(drawings, node, uriChart, part, chartXml, chartNode)
					: (ExcelChart)new ExcelPieChart(topChart, chartNode);
			case "ofPieChart":
				return topChart == null
					? new ExcelOfPieChart(drawings, node, uriChart, part, chartXml, chartNode)
					: new ExcelBarChart(topChart, chartNode);
			case "lineChart":
			case "line3DChart":
				return topChart == null
					? new ExcelLineChart(drawings, node, uriChart, part, chartXml, chartNode)
					: (ExcelChart)new ExcelLineChart(topChart, chartNode);
			case "scatterChart":
				return topChart == null
					? new ExcelScatterChart(drawings, node, uriChart, part, chartXml, chartNode)
					: (ExcelChart)new ExcelScatterChart(topChart, chartNode);
			default:
				return null;
		}
	}
	internal static ExcelChart GetNewChart(ExcelDrawings drawings, XmlNode drawNode, eChartType chartType, ExcelChart topChart, ExcelPivotTable PivotTableSource) => chartType switch
	{
		eChartType.Pie or eChartType.PieExploded or eChartType.Pie3D or eChartType.PieExploded3D => new ExcelPieChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		eChartType.BarOfPie or eChartType.PieOfPie => new ExcelOfPieChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		eChartType.Doughnut or eChartType.DoughnutExploded => new ExcelDoughnutChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		eChartType.BarClustered or eChartType.BarStacked or eChartType.BarStacked100 or eChartType.BarClustered3D or eChartType.BarStacked3D or eChartType.BarStacked1003D or eChartType.ConeBarClustered or eChartType.ConeBarStacked or eChartType.ConeBarStacked100 or eChartType.CylinderBarClustered or eChartType.CylinderBarStacked or eChartType.CylinderBarStacked100 or eChartType.PyramidBarClustered or eChartType.PyramidBarStacked or eChartType.PyramidBarStacked100 or eChartType.ColumnClustered or eChartType.ColumnStacked or eChartType.ColumnStacked100 or eChartType.Column3D or eChartType.ColumnClustered3D or eChartType.ColumnStacked3D or eChartType.ColumnStacked1003D or eChartType.ConeCol or eChartType.ConeColClustered or eChartType.ConeColStacked or eChartType.ConeColStacked100 or eChartType.CylinderCol or eChartType.CylinderColClustered or eChartType.CylinderColStacked or eChartType.CylinderColStacked100 or eChartType.PyramidCol or eChartType.PyramidColClustered or eChartType.PyramidColStacked or eChartType.PyramidColStacked100 => new ExcelBarChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		eChartType.XYScatter or eChartType.XYScatterLines or eChartType.XYScatterLinesNoMarkers or eChartType.XYScatterSmooth or eChartType.XYScatterSmoothNoMarkers => new ExcelScatterChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		eChartType.Line or eChartType.Line3D or eChartType.LineMarkers or eChartType.LineMarkersStacked or eChartType.LineMarkersStacked100 or eChartType.LineStacked or eChartType.LineStacked100 => new ExcelLineChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		eChartType.Bubble or eChartType.Bubble3DEffect => new ExcelBubbleChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		eChartType.Radar or eChartType.RadarFilled or eChartType.RadarMarkers => new ExcelRadarChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		eChartType.Surface or eChartType.SurfaceTopView or eChartType.SurfaceTopViewWireframe or eChartType.SurfaceWireframe => new ExcelSurfaceChart(drawings, drawNode, chartType, topChart, PivotTableSource),
		_ => new ExcelChart(drawings, drawNode, chartType, topChart, PivotTableSource),
	};
	public ExcelPivotTable PivotTableSource
	{
		get;
		private set;
	}
	internal void SetPivotSource(ExcelPivotTable pivotTableSource)
	{
		PivotTableSource = pivotTableSource;
		var chart = ChartXml.SelectSingleNode("c:chartSpace/c:chart", NameSpaceManager) as XmlElement;

		var pivotSource = ChartXml.CreateElement("pivotSource", ExcelPackage.schemaChart);
		chart.ParentNode.InsertBefore(pivotSource, chart);
		pivotSource.InnerXml = string.Format("<c:name>[]{0}!{1}</c:name><c:fmtId val=\"0\"/>", PivotTableSource.WorkSheet.Name, pivotTableSource.Name);

		var fmts = ChartXml.CreateElement("pivotFmts", ExcelPackage.schemaChart);
		chart.PrependChild(fmts);
		fmts.InnerXml = "<c:pivotFmt><c:idx val=\"0\"/><c:marker><c:symbol val=\"none\"/></c:marker></c:pivotFmt>";

		Series.AddPivotSerie(pivotTableSource);
	}
	internal override void DeleteMe()
	{
		try
		{
			Part.Package.DeletePart(UriChart);
		}
		catch (Exception ex)
		{
			throw (new InvalidDataException("EPPlus internal error when deleteing chart.", ex));
		}

		base.DeleteMe();
	}
}
