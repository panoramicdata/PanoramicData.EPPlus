using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Xml;

namespace PanoramicData.EPPlus.Test;

public partial class DrawingTest
{
	//[TestMethod]
	//[Ignore]
	public void Line()
	{
		var ws = _pck.Workbook.Worksheets.Add("Line");
		var chrt = ws.Drawings.AddChart("Line1", eChartType.Line) as ExcelLineChart;
		AddTestSerie(ws, chrt);
		chrt.SetSize(150);
		chrt.VaryColors = true;
		chrt.Smooth = false;
		chrt.Title.Text = "Line 3D";
		chrt.Series[0].Header = "Line serie 1";
		chrt.Axis[0].MajorGridlines.Fill.Color = Color.Black;
		chrt.Axis[0].MajorGridlines.LineStyle = eLineStyle.Dot;

		var tl = chrt.Series[0].TrendLines.Add(eTrendLine.Polynomial);
		tl.Name = "Test";
		tl.DisplayRSquaredValue = true;
		tl.DisplayEquation = true;
		tl.Forward = 15;
		tl.Backward = 1;
		tl.Intercept = 6;
		tl.Order = 5;

		chrt.Series[0].TrendLines.Add(eTrendLine.MovingAvgerage);
		chrt.Fill.Color = Color.LightSteelBlue;
		chrt.Border.LineStyle = eLineStyle.Dot;
		chrt.Border.Fill.Color = Color.Black;

		chrt.Legend.Font.Color = Color.Red;
		chrt.Legend.Font.Strike = eStrikeType.Double;
		chrt.Title.Font.Color = Color.DarkGoldenrod;
		chrt.Title.Font.LatinFont = "Arial";
		chrt.Title.Font.Bold = true;
		chrt.Title.Fill.Color = Color.White;
		chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
		chrt.Title.Border.LineStyle = eLineStyle.LongDashDotDot;
		chrt.Title.Border.Fill.Color = Color.Tomato;
		chrt.DataLabel.ShowSeriesName = true;
		chrt.DataLabel.ShowLeaderLines = true;
		chrt.EditAs = eEditAs.OneCell;
		chrt.DisplayBlanksAs = eDisplayBlanksAs.Span;
		chrt.Axis[0].Title.Text = "Axis 0";
		chrt.Axis[0].Title.Rotation = 90;
		chrt.Axis[0].Title.Overlay = true;
		chrt.Axis[1].Title.Text = "Axis 1";
		chrt.Axis[1].Title.AnchorCtr = true;
		chrt.Axis[1].Title.TextVertical = eTextVerticalType.Vertical270;
		chrt.Axis[1].Title.Border.LineStyle = eLineStyle.LongDashDotDot;

	}
	//[TestMethod]
	//[Ignore]
	public void LineMarker()
	{
		var ws = _pck.Workbook.Worksheets.Add("LineMarker1");
		var chrt = ws.Drawings.AddChart("Line1", eChartType.LineMarkers) as ExcelLineChart;
		AddTestSerie(ws, chrt);
		chrt.SetSize(150);
		chrt.Title.Text = "Line Markers";
		chrt.Series[0].Header = "Line serie 1";
		((ExcelLineChartSerie)chrt.Series[0]).Marker = eMarkerStyle.Plus;

		var chrt2 = ws.Drawings.AddChart("Line2", eChartType.LineMarkers) as ExcelLineChart;
		AddTestSerie(ws, chrt2);
		chrt2.SetPosition(500, 0);
		chrt2.SetSize(150);
		chrt2.Title.Text = "Line Markers";
		var serie = (ExcelLineChartSerie)chrt2.Series[0];
		serie.Marker = eMarkerStyle.X;

	}
	//[TestMethod]
	//[Ignore]
	public void Line2Test()
	{
		var worksheet = _pck.Workbook.Worksheets.Add("LineIssue");

		var chart = worksheet.Drawings.AddChart("LineChart", eChartType.Line);

		worksheet.Cells["A1"].Value = 1;
		worksheet.Cells["A2"].Value = 2;
		worksheet.Cells["A3"].Value = 3;
		worksheet.Cells["A4"].Value = 4;
		worksheet.Cells["A5"].Value = 5;
		worksheet.Cells["A6"].Value = 6;

		worksheet.Cells["B1"].Value = 10000;
		worksheet.Cells["B2"].Value = 10100;
		worksheet.Cells["B3"].Value = 10200;
		worksheet.Cells["B4"].Value = 10150;
		worksheet.Cells["B5"].Value = 10250;
		worksheet.Cells["B6"].Value = 10200;

		chart.Series.Add(ExcelCellBase.GetAddress(1, 2, worksheet.Dimension.End.Row, 2),
						 ExcelCellBase.GetAddress(1, 1, worksheet.Dimension.End.Row, 1));

		_ = chart.Series[0];
		chart.Axis[0].MinorGridlines.Fill.Color = Color.Red;
		chart.Axis[0].MinorGridlines.LineStyle = eLineStyle.LongDashDot;

		chart.Series[0].Header = "Blah";
	}
	//[TestMethod]
	//[Ignore]
	public void MultiChartSeries()
	{
		var worksheet = _pck.Workbook.Worksheets.Add("MultiChartTypes");

		var chart = worksheet.Drawings.AddChart("chtPie", eChartType.LineMarkers);
		chart.SetPosition(100, 100);
		chart.SetSize(800, 600);
		AddTestSerie(worksheet, chart);
		chart.Series[0].Header = "Serie5";
		chart.Style = eChartStyle.Style27;
		worksheet.Cells["W19"].Value = 120;
		worksheet.Cells["W20"].Value = 122;
		worksheet.Cells["W21"].Value = 121;
		worksheet.Cells["W22"].Value = 123;
		worksheet.Cells["W23"].Value = 125;
		worksheet.Cells["W24"].Value = 124;

		worksheet.Cells["X19"].Value = 90;
		worksheet.Cells["X20"].Value = 52;
		worksheet.Cells["X21"].Value = 88;
		worksheet.Cells["X22"].Value = 75;
		worksheet.Cells["X23"].Value = 77;
		worksheet.Cells["X24"].Value = 99;

		var cs2 = chart.PlotArea.ChartTypes.Add(eChartType.ColumnClustered);
		var s = cs2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
		s.Header = "Serie4";
		cs2.YAxis.MaxValue = 300;
		cs2.YAxis.MinValue = -5.5;
		var cs3 = chart.PlotArea.ChartTypes.Add(eChartType.Line);
		s = cs3.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["U19:U24"]);
		s.Header = "Serie1";
		cs3.UseSecondaryAxis = true;

		cs3.XAxis.Deleted = false;
		cs3.XAxis.MajorUnit = 20;
		cs3.XAxis.MinorUnit = 3;

		cs3.XAxis.TickLabelPosition = eTickLabelPosition.High;
		cs3.YAxis.LogBase = 10.2;

		var chart2 = worksheet.Drawings.AddChart("scatter1", eChartType.XYScatterSmooth);
		s = chart2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
		s.Header = "Serie2";

		var c2ct2 = chart2.PlotArea.ChartTypes.Add(eChartType.XYScatterSmooth);
		s = c2ct2.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["V19:V24"]);
		s.Header = "Serie3";
		s = c2ct2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["V19:V24"]);
		s.Header = "Serie4";

		c2ct2.UseSecondaryAxis = true;
		c2ct2.XAxis.Deleted = false;
		c2ct2.XAxis.TickLabelPosition = eTickLabelPosition.High;

		var chart3 = worksheet.Drawings.AddChart("chart", eChartType.LineMarkers);
		chart3.SetPosition(300, 1000);
		var s31 = chart3.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
		s31.Header = "Serie1";

		var c3ct2 = chart3.PlotArea.ChartTypes.Add(eChartType.LineMarkers);
		var c32 = c3ct2.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["V19:V24"]);
		c3ct2.UseSecondaryAxis = true;
		c32.Header = "Serie2";

		XmlNamespaceManager ns = new(new NameTable());
		ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		var element = chart.ChartXml.SelectSingleNode("//c:plotVisOnly", ns);
		element?.ParentNode.RemoveChild(element);
	}
	[TestMethod]
	[Ignore]
	public void ReadMultiChartSeries()
	{
		ExcelPackage pck = new(new FileInfo("c:\\temp\\chartseries.xlsx"), true);

		var ws = pck.Workbook.Worksheets[1];
		var c = ws.Drawings[0] as ExcelChart;


		var p = c.PlotArea;
		p.ChartTypes[1].Series[0].Series = "S7:S15";

		var c2 = ws.Drawings.AddChart("NewChart", eChartType.ColumnClustered);
		var serie1 = c2.Series.Add("R7:R15", "Q7:Q15");
		c2.SetSize(800, 800);
		serie1.Header = "Column Clustered";

		var subChart = c2.PlotArea.ChartTypes.Add(eChartType.LineMarkers);
		var serie2 = subChart.Series.Add("S7:S15", "Q7:Q15");
		serie2.Header = "Line";

		var subChart3 = c2.PlotArea.ChartTypes.Add(eChartType.Area);
		var serie4 = subChart3.Series.Add("R7:R15", "Q7:Q15");
		serie4.Header = "Area";
		subChart3.UseSecondaryAxis = true;

		var serie5 = subChart.Series.Add("R7:R15", "Q7:Q15");
		serie5.Header = "Line 2";

		pck.SaveAs(new FileInfo("c:\\temp\\chartseriesnew.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ChartWorksheet()
	{
		_pck = new ExcelPackage();
		var wsChart = _pck.Workbook.Worksheets.AddChart("chart", eChartType.Bubble3DEffect);
		var ws = _pck.Workbook.Worksheets.Add("data");
		AddTestSerie(ws, wsChart.Chart);
		wsChart.Chart.Style = eChartStyle.Style23;
		wsChart.Chart.Title.Text = "Chart worksheet";
		wsChart.Chart.Series[0].Header = "Serie";
		_pck.SaveAs(new FileInfo(@"c:\temp\chart.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadChartWorksheet()
	{
		_pck = new ExcelPackage(new FileInfo(@"c:\temp\chart.xlsx"));
		_ = ((ExcelChartsheet)_pck.Workbook.Worksheets[1]).Chart;

		_pck.SaveAs(new FileInfo(@"c:\temp\chart.xlsx"));

	}
	[Ignore]
	[TestMethod]
	public void ReadWriteSmoothChart()
	{
		_pck = new ExcelPackage(new FileInfo(@"c:\temp\bug\Xds_2014_TEST.xlsx"));
		_ = _pck.Workbook.Worksheets[1].Drawings[0] as ExcelChart;
		_pck.Workbook.Worksheets[1].Cells["B2"].Value = 33;
		_pck.SaveAs(new FileInfo(@"c:\temp\chart.xlsx"));

	}
}
