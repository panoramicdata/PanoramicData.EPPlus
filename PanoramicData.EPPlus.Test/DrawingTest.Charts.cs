using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class DrawingTest
{
	//[TestMethod]
	// [Ignore]
	public void BarChart()
	{
		var ws = _pck.Workbook.Worksheets.Add("BarChart");
		var chrt = ws.Drawings.AddChart("barChart", eChartType.BarClustered) as ExcelBarChart;
		chrt.SetPosition(50, 50);
		chrt.SetSize(800, 300);
		AddTestSerie(ws, chrt);
		chrt.VaryColors = true;
		chrt.XAxis.Orientation = eAxisOrientation.MaxMin;
		chrt.XAxis.MajorTickMark = eAxisTickMark.In;
		chrt.XAxis.Format = "yyyy-MM";
		chrt.YAxis.Orientation = eAxisOrientation.MaxMin;
		chrt.YAxis.MinorTickMark = eAxisTickMark.Out;
		chrt.ShowHiddenData = true;
		chrt.DisplayBlanksAs = eDisplayBlanksAs.Zero;
		chrt.Title.RichText.Text = "Barchart Test";
		chrt.GapWidth = 5;
		Assert.IsTrue(chrt.ChartType == eChartType.BarClustered, "Invalid Charttype");
		Assert.IsTrue(chrt.Direction == eDirection.Bar, "Invalid Bardirection");
		Assert.IsTrue(chrt.Grouping == eGrouping.Clustered, "Invalid Grouping");
		Assert.IsTrue(chrt.Shape == eShape.Box, "Invalid Shape");
	}

	private static void AddTestSerie(ExcelWorksheet ws, ExcelChart chrt)
	{
		AddTestData(ws);
		chrt.Series.Add("'" + ws.Name + "'!V19:V24", "'" + ws.Name + "'!U19:U24");
	}

	private static void AddTestData(ExcelWorksheet ws)
	{
		ws.Cells["U19"].Value = new DateTime(2009, 12, 31);
		ws.Cells["U20"].Value = new DateTime(2010, 1, 1);
		ws.Cells["U21"].Value = new DateTime(2010, 1, 2);
		ws.Cells["U22"].Value = new DateTime(2010, 1, 3);
		ws.Cells["U23"].Value = new DateTime(2010, 1, 4);
		ws.Cells["U24"].Value = new DateTime(2010, 1, 5);
		ws.Cells["U19:U24"].Style.Numberformat.Format = "yyyy-mm-dd";

		ws.Cells["V19"].Value = 100;
		ws.Cells["V20"].Value = 102;
		ws.Cells["V21"].Value = 101;
		ws.Cells["V22"].Value = 103;
		ws.Cells["V23"].Value = 105;
		ws.Cells["V24"].Value = 104;

		ws.Cells["W19"].Value = 105;
		ws.Cells["W20"].Value = 108;
		ws.Cells["W21"].Value = 104;
		ws.Cells["W22"].Value = 121;
		ws.Cells["W23"].Value = 103;
		ws.Cells["W24"].Value = 109;


		ws.Cells["X19"].Value = "öäå";
		ws.Cells["X20"].Value = "ÖÄÅ";
		ws.Cells["X21"].Value = "üÛ";
		ws.Cells["X22"].Value = "&%#¤";
		ws.Cells["X23"].Value = "ÿ";
		ws.Cells["X24"].Value = "û";
	}
	//[TestMethod]
	//[Ignore]
	public void PieChart()
	{
		var ws = _pck.Workbook.Worksheets.Add("PieChart");
		var chrt = ws.Drawings.AddChart("pieChart", eChartType.Pie) as ExcelPieChart;

		AddTestSerie(ws, chrt);

		chrt.To.Row = 25;
		chrt.To.Column = 12;

		chrt.DataLabel.ShowPercent = true;
		chrt.Legend.Font.Color = Color.SteelBlue;
		chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
		chrt.Legend.Position = eLegendPosition.TopRight;
		Assert.IsTrue(chrt.ChartType == eChartType.Pie, "Invalid Charttype");
		Assert.IsTrue(chrt.VaryColors);
		chrt.Title.Text = "Piechart";
	}
	public void PieOfChart()
	{
		var ws = _pck.Workbook.Worksheets.Add("PieOfChart");
		var chrt = ws.Drawings.AddChart("pieOfChart", eChartType.BarOfPie) as ExcelOfPieChart;

		AddTestSerie(ws, chrt);

		chrt.To.Row = 25;
		chrt.To.Column = 12;

		chrt.DataLabel.ShowPercent = true;
		chrt.Legend.Font.Color = Color.SteelBlue;
		chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
		chrt.Legend.Position = eLegendPosition.TopRight;
		Assert.IsTrue(chrt.ChartType == eChartType.BarOfPie, "Invalid Charttype");
		chrt.Title.Text = "Piechart";
	}
	//[TestMethod]
	//[Ignore]
	public void PieChart3D()
	{
		var ws = _pck.Workbook.Worksheets.Add("PieChart3d");
		var chrt = ws.Drawings.AddChart("pieChart3d", eChartType.Pie3D) as ExcelPieChart;
		AddTestSerie(ws, chrt);

		chrt.To.Row = 25;
		chrt.To.Column = 12;

		chrt.DataLabel.ShowValue = true;
		chrt.Legend.Position = eLegendPosition.Left;
		chrt.ShowHiddenData = false;
		chrt.DisplayBlanksAs = eDisplayBlanksAs.Gap;
		chrt.Title.RichText.Add("Pie RT Title add");
		Assert.IsTrue(chrt.ChartType == eChartType.Pie3D, "Invalid Charttype");
		Assert.IsTrue(chrt.VaryColors);

	}
	//[TestMethod]
	//[Ignore]
	public void Scatter()
	{
		var ws = _pck.Workbook.Worksheets.Add("Scatter");
		var chrt = ws.Drawings.AddChart("ScatterChart1", eChartType.XYScatterSmoothNoMarkers) as ExcelScatterChart;
		AddTestSerie(ws, chrt);
		chrt.To.Row = 23;
		chrt.To.Column = 12;
		var r1 = chrt.Title.RichText.Add("Header");
		r1.Bold = true;
		var r2 = chrt.Title.RichText.Add("  Text");
		r2.UnderLine = eUnderLineType.WavyHeavy;

		chrt.Title.Fill.Style = eFillStyle.SolidFill;
		chrt.Title.Fill.Color = Color.LightBlue;
		chrt.Title.Fill.Transparancy = 50;
		chrt.VaryColors = true;
		var ser = chrt.Series[0] as ExcelScatterChartSerie;
		ser.DataLabel.Position = eLabelPosition.Center;
		ser.DataLabel.ShowValue = true;
		ser.DataLabel.ShowCategory = true;
		ser.DataLabel.Fill.Color = Color.BlueViolet;
		ser.DataLabel.Font.Color = Color.White;
		ser.DataLabel.Font.Italic = true;
		ser.DataLabel.Font.LatinFont = "bookman old style";
		ser.DataLabel.Font.ComplexFont = "bookman old style";
		ser.DataLabel.Font.Size = 8;
		Assert.IsTrue(chrt.ChartType == eChartType.XYScatterSmoothNoMarkers, "Invalid Charttype");
		chrt.Series[0].Header = "Test serie";
		chrt = ws.Drawings.AddChart("ScatterChart2", eChartType.XYScatterSmooth) as ExcelScatterChart;
		chrt.Series.Add("U19:U24", "V19:V24");

		chrt.From.Column = 0;
		chrt.From.Row = 25;
		chrt.To.Row = 53;
		chrt.To.Column = 12;
		chrt.Legend.Position = eLegendPosition.Bottom;
	}
	//[TestMethod]
	//[Ignore]
	public void Bubble()
	{
		var ws = _pck.Workbook.Worksheets.Add("Bubble");
		var chrt = ws.Drawings.AddChart("Bubble", eChartType.Bubble) as ExcelBubbleChart;
		AddTestData(ws);

		chrt.Series.Add("V19:V24", "U19:U24");

		chrt = ws.Drawings.AddChart("Bubble3d", eChartType.Bubble3DEffect) as ExcelBubbleChart;
		ws.Cells["W19"].Value = 1;
		ws.Cells["W20"].Value = 1;
		ws.Cells["W21"].Value = 2;
		ws.Cells["W22"].Value = 2;
		ws.Cells["W23"].Value = 3;
		ws.Cells["W24"].Value = 4;

		chrt.Series.Add("V19:V24", "U19:U24", "W19:W24");
		chrt.Style = eChartStyle.Style25;

		chrt.From.Row = 23;
		chrt.From.Column = 12;
		chrt.To.Row = 33;
		chrt.To.Column = 22;
		chrt.Title.Text = "Header Text";


	}
	//[TestMethod]
	//[Ignore]
	public void Radar()
	{
		var ws = _pck.Workbook.Worksheets.Add("Radar");
		AddTestData(ws);

		var chrt = ws.Drawings.AddChart("Radar1", eChartType.Radar) as ExcelRadarChart;
		var s = chrt.Series.Add("V19:V24", "U19:U24");
		s.Header = "serie1";
		chrt.From.Row = 23;
		chrt.From.Column = 12;
		chrt.To.Row = 38;
		chrt.To.Column = 22;
		chrt.Title.Text = "Radar Chart 1";

		chrt = ws.Drawings.AddChart("Radar2", eChartType.RadarFilled) as ExcelRadarChart;
		s = chrt.Series.Add("V19:V24", "U19:U24");
		s.Header = "serie1";
		chrt.From.Row = 43;
		chrt.From.Column = 12;
		chrt.To.Row = 58;
		chrt.To.Column = 22;
		chrt.Title.Text = "Radar Chart 2";

		chrt = ws.Drawings.AddChart("Radar3", eChartType.RadarMarkers) as ExcelRadarChart;
		var rs = (ExcelRadarChartSerie)chrt.Series.Add("V19:V24", "U19:U24");
		rs.Header = "serie1";
		rs.Marker = eMarkerStyle.Star;
		rs.MarkerSize = 14;

		chrt.From.Row = 63;
		chrt.From.Column = 12;
		chrt.To.Row = 78;
		chrt.To.Column = 22;
		chrt.Title.Text = "Radar Chart 3";
	}
	//[TestMethod]
	//[Ignore]
	public void Surface()
	{
		var ws = _pck.Workbook.Worksheets.Add("Surface");
		AddTestData(ws);

		var chrt = ws.Drawings.AddChart("Surface1", eChartType.Surface) as ExcelSurfaceChart;
		var s = chrt.Series.Add("V19:V24", "U19:U24");
		chrt.Series.Add("W19:W24", "U19:U24");
		s.Header = "serie1";
		chrt.From.Row = 23;
		chrt.From.Column = 12;
		chrt.To.Row = 38;
		chrt.To.Column = 22;
		chrt.Title.Text = "Surface Chart 1";
	}
	//[TestMethod]
	//[Ignore]
	public void Pyramid()
	{
		var ws = _pck.Workbook.Worksheets.Add("Pyramid");
		var chrt = ws.Drawings.AddChart("Pyramid1", eChartType.PyramidCol) as ExcelBarChart;
		AddTestSerie(ws, chrt);
		chrt.VaryColors = true;
		chrt.To.Row = 23;
		chrt.To.Column = 12;
		chrt.Title.Text = "Header Text";
		chrt.Title.Fill.Style = eFillStyle.SolidFill;
		chrt.Title.Fill.Color = Color.DarkBlue;
		chrt.DataLabel.ShowValue = true;
		chrt.Border.LineCap = eLineCap.Round;
		chrt.Border.LineStyle = eLineStyle.LongDashDotDot;
		chrt.Border.Fill.Style = eFillStyle.SolidFill;
		chrt.Border.Fill.Color = Color.Blue;

		chrt.Fill.Color = Color.LightCyan;
		chrt.PlotArea.Fill.Color = Color.White;
		chrt.PlotArea.Border.Fill.Style = eFillStyle.SolidFill;
		chrt.PlotArea.Border.Fill.Color = Color.Beige;
		chrt.PlotArea.Border.LineStyle = eLineStyle.LongDash;

		chrt.Legend.Fill.Color = Color.Aquamarine;
		chrt.Legend.Position = eLegendPosition.Top;
		chrt.Axis[0].Fill.Style = eFillStyle.SolidFill;
		chrt.Axis[0].Fill.Color = Color.Black;
		chrt.Axis[0].Font.Color = Color.White;

		chrt.Axis[1].Fill.Style = eFillStyle.SolidFill;
		chrt.Axis[1].Fill.Color = Color.LightSlateGray;
		chrt.Axis[1].Font.Color = Color.DarkRed;

		chrt.DataLabel.Font.Bold = true;
		chrt.DataLabel.Fill.Color = Color.LightBlue;
		chrt.DataLabel.Border.Fill.Style = eFillStyle.SolidFill;
		chrt.DataLabel.Border.Fill.Color = Color.Black;
		chrt.DataLabel.Border.LineStyle = eLineStyle.Solid;
	}
	//[TestMethod]
	//[Ignore]
	public void Cone()
	{
		var ws = _pck.Workbook.Worksheets.Add("Cone");
		var chrt = ws.Drawings.AddChart("Cone1", eChartType.ConeBarClustered) as ExcelBarChart;
		AddTestSerie(ws, chrt);
		chrt.VaryColors = true;
		chrt.SetSize(200);
		chrt.Title.Text = "Cone bar";
		chrt.Series[0].Header = "Serie 1";
		chrt.Legend.Position = eLegendPosition.Right;
		chrt.Axis[1].DisplayUnit = 100000;
		Assert.AreEqual(100000, chrt.Axis[1].DisplayUnit);
	}
	//[TestMethod]
	//[Ignore]
	public void Column()
	{
		var ws = _pck.Workbook.Worksheets.Add("Column");
		var chrt = ws.Drawings.AddChart("Column1", eChartType.ColumnClustered3D) as ExcelBarChart;
		AddTestSerie(ws, chrt);
		chrt.VaryColors = true;
		chrt.View3D.RightAngleAxes = true;
		chrt.View3D.DepthPercent = 99;
		chrt.View3D.RightAngleAxes = true;
		chrt.SetSize(200);
		chrt.Title.Text = "Column";
		chrt.Series[0].Header = "Serie 1";
		chrt.Locked = false;
		chrt.Print = false;
		chrt.EditAs = eEditAs.TwoCell;
		chrt.Axis[1].DisplayUnit = 10020;
		Assert.AreEqual(10020, chrt.Axis[1].DisplayUnit);
	}
	//[TestMethod]
	//[Ignore]
	public void Dougnut()
	{
		var ws = _pck.Workbook.Worksheets.Add("Dougnut");
		var chrt = ws.Drawings.AddChart("Dougnut1", eChartType.DoughnutExploded) as ExcelDoughnutChart;
		AddTestSerie(ws, chrt);
		chrt.SetSize(200);
		chrt.Title.Text = "Doughnut Exploded";
		chrt.Series[0].Header = "Serie 1";
		chrt.EditAs = eEditAs.Absolute;
	}
}
