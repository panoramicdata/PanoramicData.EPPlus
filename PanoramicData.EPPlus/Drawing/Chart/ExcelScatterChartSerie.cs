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
 * Jan Källman		Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A serie for a scatter chart
/// </summary>
public sealed class ExcelScatterChartSerie : ExcelChartSerie
{
	/// <summary>
	/// Default constructor
	/// </summary>
	/// <param name="chartSeries">Parent collection</param>
	/// <param name="ns">Namespacemanager</param>
	/// <param name="node">Topnode</param>
	/// <param name="isPivot">Is pivotchart</param>
	internal ExcelScatterChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
		base(chartSeries, ns, node, isPivot)
	{
		if (chartSeries.Chart.ChartType is eChartType.XYScatterLines or
			eChartType.XYScatterSmooth)
		{
			Marker = eMarkerStyle.Square;
		}

		if (chartSeries.Chart.ChartType is eChartType.XYScatterSmooth or
			eChartType.XYScatterSmoothNoMarkers)
		{
			Smooth = 1;
		}
		else if (chartSeries.Chart.ChartType is eChartType.XYScatterLines or eChartType.XYScatterLinesNoMarkers or eChartType.XYScatter)

		{
			Smooth = 0;
		}
	}
	ExcelChartSerieDataLabel _DataLabel;
	/// <summary>
	/// Datalabel
	/// </summary>
	public ExcelChartSerieDataLabel DataLabel => GetDataLabel(ref _DataLabel);
	const string smoothPath = "c:smooth/@val";
	/// <summary>
	/// Smooth for scattercharts
	/// </summary>
	public int Smooth
	{
		get
		{
			return GetXmlNodeInt(smoothPath);
		}
		internal set
		{
			SetXmlNodeString(smoothPath, value.ToString());
		}
	}
	/// <summary>
	/// Marker symbol 
	/// </summary>
	public eMarkerStyle Marker
	{
		get
		{
			return GetMarker();
		}
		set
		{
			SetMarker(value);
		}
	}

	//new properties for excel scatter-plots: LineColor, MarkerSize, MarkerColor, LineWidth and MarkerLineColor
	//implemented according to https://epplus.codeplex.com/discussions/287917
	readonly string LINECOLOR_PATH = "c:spPr/a:ln/a:solidFill/a:srgbClr/@val";
	/// <summary>
	/// Line color.
	/// </summary>
	///
	/// <value>
	/// The color of the line.
	/// </value>
	public Color LineColor
	{
		get
		{
			return GetColorWithAlpha(LINECOLOR_PATH);
		}
		set
		{
			SetColorWithAlpha(LINECOLOR_PATH, value);
		}
	}
	/// <summary>
	/// Gets or sets the size of the marker.
	/// </summary>
	///
	/// <remarks>
	/// value between 2 and 72.
	/// </remarks>
	///
	/// <value>
	/// The size of the marker.
	/// </value>
	public int MarkerSize
	{
		get
		{
			return GetMarkerSize();
		}
		set
		{
			SetMarkerSize(value);
		}
	}
	readonly string MARKERCOLOR_PATH = "c:marker/c:spPr/a:solidFill/a:srgbClr/@val";
	/// <summary>
	/// Marker color.
	/// </summary>
	///
	/// <value>
	/// The color of the Marker.
	/// </value>
	public Color MarkerColor
	{
		get
		{
			return GetColorWithAlpha(MARKERCOLOR_PATH);
		}
		set
		{
			SetColorWithAlpha(MARKERCOLOR_PATH, value);
		}
	}

	/// <summary>
	/// Gets or sets the width of the line in pt.
	/// </summary>
	///
	/// <value>
	/// The width of the line.
	/// </value>
	public double LineWidth
	{
		get
		{
			return GetLineWidth();
		}
		set
		{
			SetLineWidth(value);
		}
	}
	//marker line color
	readonly string MARKERLINECOLOR_PATH = "c:marker/c:spPr/a:ln/a:solidFill/a:srgbClr/@val";
	/// <summary>
	/// Marker Line color.
	/// (not to be confused with LineColor)
	/// </summary>
	///
	/// <value>
	/// The color of the Marker line.
	/// </value>
	public Color MarkerLineColor
	{
		get
		{
			return GetColorWithAlpha(MARKERLINECOLOR_PATH);
		}
		set
		{
			SetColorWithAlpha(MARKERLINECOLOR_PATH, value);
		}
	}


	private Color GetColorWithAlpha(string path)
	{
		var color = GetXmlNodeString(path);
		if (color == "")
		{
			return Color.Black;
		}

		var c = Color.FromArgb(Convert.ToInt32(color, 16));
		var a = getAlphaChannel(path);
		return a != 255 ? Color.FromArgb(a, c) : c;
	}

	// The alpha byte is cut from the hex value and written to its own node
	private void SetColorWithAlpha(string path, Color value)
	{
		SetXmlNodeString(path, value.ToArgb().ToString("X8")[2..], true);
		setAlphaChannel(value, path);
	}

	/// <summary>
	/// write alpha value (if Color.A != 255)
	/// </summary>
	/// <param name="c">Color</param>
	/// <param name="xPath">where to write</param>
	/// <remarks>
	/// alpha-values may only written to color-nodes
	/// eg: a:prstClr (preset), a:hslClr (hsl), a:schemeClr (schema), a:sysClr (system), a:scrgbClr (rgb percent) or a:srgbClr (rgb hex)
	///     .../a:prstClr/a:alpha/@val
	/// </remarks>
	private void setAlphaChannel(Color c, string xPath)
	{
		//check 4 Alpha-values
		if (c.A != 255)
		{ //opaque color => alpha == 255 //source: https://msdn.microsoft.com/en-us/library/1hstcth9%28v=vs.110%29.aspx
		  //check path
			var s = xPath4Alpha(xPath);
			if (s.Length > 0)
			{
				var alpha = ((c.A == 0) ? 0 : (100 - c.A) * 1000).ToString(); //note: excel writes 100% transparency (alpha=0) as "0" and not as "100000"
				SetXmlNodeString(s, alpha, true);
			}
		}
	}
	/// <summary>
	/// read AlphaChannel from a:solidFill
	/// </summary>
	/// <param name="xPath"></param>
	/// <returns>alpha or 255 if their is no such node</returns>
	private int getAlphaChannel(string xPath)
	{
		var r = 255;
		var s = xPath4Alpha(xPath);
		if (s.Length > 0)
		{
			if (int.TryParse(GetXmlNodeString(s), NumberStyles.Any, CultureInfo.InvariantCulture, out var i))
			{
				r = (i == 0) ? 0 : 100 - (i / 1000);
			}
		}

		return r;
	}
	/// <summary>
	/// creates xPath to alpha attribute for a color 
	/// eg: a:prstClr/a:alpha/@val
	/// </summary>
	/// <param name="xPath">xPath to color node</param>
	/// <returns></returns>
	private static string xPath4Alpha(string xPath)
	{
		var s = string.Empty;
		if (xPath.EndsWith("@val", StringComparison.Ordinal))
		{
			xPath = xPath[..xPath.IndexOf("@val", StringComparison.Ordinal)];
		}

		if (xPath.EndsWith('/'))
		{ //cut tailing slash
			xPath = xPath[..^1];
		}
		//parent node must be a color node/definition
		List<string> colorDefs = ["a:prstClr", "a:hslClr", "a:schemeClr", "a:sysClr", "a:scrgbClr", "a:srgbClr"];
		if (colorDefs.Find(cd => xPath.EndsWith(cd, StringComparison.Ordinal)) != null)
		{
			s = xPath + "/a:alpha/@val";
		}
		else
		{
			System.Diagnostics.Debug.Assert(false);
			throw new InvalidOperationException("alpha-values can only set to Colors");
		}

		return s;
	}
}
