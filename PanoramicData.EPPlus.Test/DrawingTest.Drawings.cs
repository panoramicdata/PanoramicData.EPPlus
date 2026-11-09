using EPPlusTest.Properties;
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
	public void ReadDrawing()
	{
		using ExcelPackage pck = new(new FileInfo(_worksheetPath + @"Drawing.xlsx"));
		var ws = pck.Workbook.Worksheets["Pyramid"];
		Assert.AreEqual(104D, ws.Cells["V24"].Value);
		ws = pck.Workbook.Worksheets["Scatter"];
		var cht = ws.Drawings["ScatterChart1"] as ExcelScatterChart;
		Assert.AreEqual("Header  Text", cht.Title.Text);
		cht.Title.Text = "Test";
		Assert.AreEqual("Test", cht.Title.Text);
	}

	public void Picture()
	{
		var ws = _pck.Workbook.Worksheets.Add("Picture");
		ws.Drawings.AddPicture("Pic1", Resources.Test1);

		var pic = ws.Drawings.AddPicture("Pic2", Resources.Test1);
		pic.SetPosition(150, 200);
		pic.Border.LineStyle = eLineStyle.Solid;
		pic.Border.Fill.Color = Color.DarkCyan;
		pic.Fill.Style = eFillStyle.SolidFill;
		pic.Fill.Color = Color.White;
		pic.Fill.Transparancy = 50;

		pic = ws.Drawings.AddPicture("Pic3", Resources.Test1);
		pic.SetPosition(400, 200);
		pic.SetSize(150);

		pic = ws.Drawings.AddPicture("Pic5", new FileInfo(Path.Combine(_clipartPath, "BitmapImage.gif")));
		pic.SetPosition(400, 200);
		pic.SetSize(150);

		ws.Column(1).Width = 53;
		ws.Column(4).Width = 58;

		pic = ws.Drawings.AddPicture("Pic6öäå", new FileInfo(Path.Combine(_clipartPath, "BitmapImage.gif")));
		pic.SetPosition(400, 400);
		pic.SetSize(100);

		pic = ws.Drawings.AddPicture("PicPixelSized", Resources.Test1);
		pic.SetPosition(800, 800);
		pic.SetSize(568 * 2, 66 * 2);
		var ws2 = _pck.Workbook.Worksheets.Add("Picture2");
		var fi = new FileInfo(Path.Combine(_clipartPath, "BitmapImage.gif"));
		if (fi.Exists)
		{
			ws2.Drawings.AddPicture("Pic7", fi);
		}
		else
		{
#if (!Core)
                TestContext.WriteLine("AG00021_.GIF does not exists. Skipping Pic7.");
#endif
		}

		_pck.Workbook.Worksheets.Add("Picture3", ws2);
	}
	//[TestMethod]
	//[Ignore]
	public void DrawingSizingAndPositioning()
	{
		var ws = _pck.Workbook.Worksheets.Add("DrawingPosSize");

		var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);
		pic.SetPosition(1, 0, 1, 0);

		pic = ws.Drawings.AddPicture("Pic2", Resources.Test1);
		pic.EditAs = eEditAs.Absolute;
		pic.SetPosition(10, 5, 1, 4);

		pic = ws.Drawings.AddPicture("Pic3", Resources.Test1);
		pic.EditAs = eEditAs.TwoCell;
		pic.SetPosition(20, 5, 2, 4);


		ws.Column(1).Width = 100;
		ws.Column(3).Width = 100;
	}
	//[TestMethod]
	//[Ignore]
	public void Drawings()
	{
		var ws = _pck.Workbook.Worksheets.Add("Shapes");
		int y = 100, i = 1;
		foreach (var style in Enum.GetValues<eShapeStyle>())
		{
			var shape = ws.Drawings.AddShape("shape" + i.ToString(), style);
			shape.SetPosition(y, 100);
			shape.SetSize(300, 300);
			y += 400;
			shape.Text = style.ToString();
			i++;
		}

		(ws.Drawings["shape1"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;
		var rt = (ws.Drawings["shape1"] as ExcelShape).RichText.Add("Added formated richtext");
		(ws.Drawings["shape1"] as ExcelShape).LockText = false;
		rt.Bold = true;
		rt.Color = Color.Aquamarine;
		rt.Italic = true;
		rt.Size = 17;
		(ws.Drawings["shape2"] as ExcelShape).TextVertical = eTextVerticalType.Vertical;
		rt = (ws.Drawings["shape2"] as ExcelShape).RichText.Add("\r\nAdded formated richtext");
		rt.Bold = true;
		rt.Color = Color.DarkGoldenrod;
		rt.LatinFont = "Times new roman";
		rt.ComplexFont = "Times new roman";
		rt.Size = 18;
		rt.UnderLine = eUnderLineType.Single;
		rt.UnderLineColor = Color.Green;


		(ws.Drawings["shape3"] as ExcelShape).TextAnchoring = eTextAnchoringType.Bottom;
		(ws.Drawings["shape3"] as ExcelShape).TextAnchoringControl = true;

		(ws.Drawings["shape4"] as ExcelShape).TextVertical = eTextVerticalType.Vertical270;
		(ws.Drawings["shape4"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;

		(ws.Drawings["shape5"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
		(ws.Drawings["shape5"] as ExcelShape).Fill.Color = Color.Red;
		(ws.Drawings["shape5"] as ExcelShape).Fill.Transparancy = 50;

		(ws.Drawings["shape6"] as ExcelShape).Fill.Style = eFillStyle.NoFill;
		(ws.Drawings["shape6"] as ExcelShape).Font.Color = Color.Black;
		(ws.Drawings["shape6"] as ExcelShape).Border.Fill.Color = Color.Black;

		(ws.Drawings["shape7"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
		(ws.Drawings["shape7"] as ExcelShape).Fill.Color = Color.Gray;
		(ws.Drawings["shape7"] as ExcelShape).Border.Fill.Style = eFillStyle.SolidFill;
		(ws.Drawings["shape7"] as ExcelShape).Border.Fill.Color = Color.Black;
		(ws.Drawings["shape7"] as ExcelShape).Border.Fill.Transparancy = 43;
		(ws.Drawings["shape7"] as ExcelShape).Border.LineCap = eLineCap.Round;
		(ws.Drawings["shape7"] as ExcelShape).Border.LineStyle = eLineStyle.LongDash;
		(ws.Drawings["shape7"] as ExcelShape).Font.UnderLineColor = Color.Blue;
		(ws.Drawings["shape7"] as ExcelShape).Font.Color = Color.Black;
		(ws.Drawings["shape7"] as ExcelShape).Font.Bold = true;
		(ws.Drawings["shape7"] as ExcelShape).Font.LatinFont = "Arial";
		(ws.Drawings["shape7"] as ExcelShape).Font.ComplexFont = "Arial";
		(ws.Drawings["shape7"] as ExcelShape).Font.Italic = true;
		(ws.Drawings["shape7"] as ExcelShape).Font.UnderLine = eUnderLineType.Dotted;

		(ws.Drawings["shape8"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
		(ws.Drawings["shape8"] as ExcelShape).Font.LatinFont = "Miriam";
		(ws.Drawings["shape8"] as ExcelShape).Font.UnderLineColor = Color.CadetBlue;
		(ws.Drawings["shape8"] as ExcelShape).Font.UnderLine = eUnderLineType.Single;

		(ws.Drawings["shape9"] as ExcelShape).TextAlignment = eTextAlignment.Right;

		(ws.Drawings["shape120"] as ExcelShape).LineEnds.TailEnd = eEndStyle.Oval;
		(ws.Drawings["shape120"] as ExcelShape).LineEnds.TailEndSizeWidth = eEndSize.Large;
		(ws.Drawings["shape120"] as ExcelShape).LineEnds.TailEndSizeHeight = eEndSize.Large;
		(ws.Drawings["shape120"] as ExcelShape).LineEnds.HeadEnd = eEndStyle.Arrow;
		(ws.Drawings["shape120"] as ExcelShape).LineEnds.HeadEndSizeHeight = eEndSize.Small;
		(ws.Drawings["shape120"] as ExcelShape).LineEnds.HeadEndSizeWidth = eEndSize.Small;
	}
	[TestMethod]
	[Ignore]
	public void DrawingWorksheetCopy()
	{
		_pck.Workbook.Worksheets.Add("Copy Shapes", _pck.Workbook.Worksheets["Shapes"]);
		_pck.Workbook.Worksheets.Add("Copy Scatter", _pck.Workbook.Worksheets["Scatter"]);
		_pck.Workbook.Worksheets.Add("Copy Picture", _pck.Workbook.Worksheets["Picture"]);
	}
	//[TestMethod]
	//[Ignore]
	public void DeleteDrawing()
	{
		var ws = _pck.Workbook.Worksheets.Add("DeleteDrawing1");
		ws.Drawings.AddChart("Chart1", eChartType.Line);
		var chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
		ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
		ws.Drawings.AddPicture("Pic1", Resources.Test1);
		ws.Drawings.Remove(2);
		ws.Drawings.Remove(chart2);
		ws.Drawings.Remove("Pic1");

		ws = _pck.Workbook.Worksheets.Add("DeleteDrawing2");
		ws.Drawings.AddChart("Chart1", eChartType.Line);
		ws.Drawings.AddChart("Chart2", eChartType.Line);
		ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
		ws.Drawings.AddPicture("Pic1", Resources.Test1);

		ws.Drawings.Remove("chart1");

		ws = _pck.Workbook.Worksheets.Add("ClearDrawing2");
		ws.Drawings.AddChart("Chart1", eChartType.Line);
		ws.Drawings.AddChart("Chart2", eChartType.Line);
		ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
		ws.Drawings.AddPicture("Pic1", Resources.Test1);
		ws.Drawings.Clear();
	}
	//[TestMethod]
	//[Ignore]
	public void ReadDocument()
	{
		var fi = new FileInfo(_worksheetPath + "Drawing.xlsx");
		if (!fi.Exists)
		{
			Assert.Inconclusive("Drawing.xlsx is not created. Skippng");
		}

		var pck = new ExcelPackage(fi, true);

		foreach (var ws in pck.Workbook.Worksheets)
		{
			foreach (ExcelDrawing d in pck.Workbook.Worksheets[1].Drawings)
			{
				if (d is ExcelChart chart)
				{
#if (!Core)
                        TestContext.WriteLine(chart.ChartType.ToString());
#endif
				}
			}
		}

		pck.Dispose();
	}
	[TestMethod]
	public void TestHeaderaddress()
	{
		_pck = new ExcelPackage();
		var ws = _pck.Workbook.Worksheets.Add("Draw");
		var chart = ws.Drawings.AddChart("NewChart1", eChartType.Area) as ExcelChart;
		var ser1 = chart.Series.Add("A1:A2", "B1:B2");
		ser1.HeaderAddress = new ExcelAddress("A1:A2");
		ser1.HeaderAddress = new ExcelAddress("A1:B1");
		ser1.HeaderAddress = new ExcelAddress("A1");
		_pck.Dispose();
		_pck = null;
	}
	[Ignore]
	[TestMethod]
	public void AllDrawingsInsideMarkupCompatibility()
	{
		var workbooksDir = Path.Combine(AppContext.BaseDirectory, Path.Combine("..", "..", "workbooks"));

		// This is codeplex issue 15028: Making an unrelated change to an Excel file that contains drawings that ALL exist
		// inside MarkupCompatibility/Choice nodes causes the drawings.xml file to be incorrectly garbage collected
		// when an unrelated change is made.
		var path = Path.Combine(workbooksDir, "AllDrawingsInsideMarkupCompatibility.xlsm");

		// Load example document.
		_pck = new ExcelPackage(new FileInfo(path));
		// Verify the drawing part exists:
		Uri partUri = new("/xl/drawings/drawing1.xml", UriKind.Relative);
		Assert.IsTrue(_pck.Package.PartExists(partUri));

		// The Excel Drawings NamespaceManager from ExcelDrawing.CreateNSM:
		NameTable nt = new();
		var xmlNsm = new XmlNamespaceManager(nt);
		xmlNsm.AddNamespace("a", ExcelPackage.schemaDrawings);
		xmlNsm.AddNamespace("xdr", ExcelPackage.schemaSheetDrawings);
		xmlNsm.AddNamespace("c", ExcelPackage.schemaChart);
		xmlNsm.AddNamespace("r", ExcelPackage.schemaRelationships);
		xmlNsm.AddNamespace("mc", ExcelPackage.schemaMarkupCompatibility);

		XmlDocument drawingsXml = new()
		{
			PreserveWhitespace = false
		};
		XmlHelper.LoadXmlSafe(drawingsXml, _pck.Package.GetPart(partUri).GetStream());

		// Verify that there are the correct # of drawings:
		Assert.AreEqual(5, drawingsXml.SelectNodes("//*[self::xdr:twoCellAnchor or self::xdr:oneCellAnchor or self::xdr:absoluteAnchor]", xmlNsm).Count);

		// Make unrelated change. (in this case a useless additional worksheet)
		_pck.Workbook.Worksheets.Add("NewWorksheet");

		// Save it out.
		var savedPath = Path.Combine(workbooksDir, "AllDrawingsInsideMarkupCompatibility2.xlsm");
		_pck.SaveAs(new FileInfo(savedPath));
		_pck.Dispose();

		// Reload the new saved file.
		_pck = new ExcelPackage(new FileInfo(savedPath));

		// Verify the drawing part still exists.
		Assert.IsTrue(_pck.Package.PartExists(new Uri("/xl/drawings/drawing1.xml", UriKind.Relative)));

		drawingsXml = new XmlDocument
		{
			PreserveWhitespace = false
		};
		XmlHelper.LoadXmlSafe(drawingsXml, _pck.Package.GetPart(partUri).GetStream());

		// Verify that there are the correct # of drawings:
		Assert.AreEqual(5, drawingsXml.SelectNodes("//*[self::xdr:twoCellAnchor or self::xdr:oneCellAnchor or self::xdr:absoluteAnchor]", xmlNsm).Count);
		// Verify that the new worksheet exists:
		Assert.IsNotNull(_pck.Workbook.Worksheets["NewWorksheet"]);
		// Cleanup:
		File.Delete(savedPath);
	}
	public void DrawingRowheightDynamic()
	{
		var ws = _pck.Workbook.Worksheets.Add("PicResize");
		ws.Cells["A1"].Value = "test";
		ws.Cells["A1"].Style.Font.Name = "Symbol";
		ws.Cells["A1"].Style.Font.Size = 39;
		ws.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Symbol";
		ws.Workbook.Styles.NamedStyles[0].Style.Font.Size = 16;
		var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);
		pic.SetPosition(10, 12);
	}
	[TestMethod]
	public void DrawingWidthAdjust()
	{
		// Needs a local customer template that is not in the repository.
	}
}
