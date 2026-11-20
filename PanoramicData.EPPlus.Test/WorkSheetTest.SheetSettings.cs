using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	//[Ignore]
	//[TestMethod]
	public void WorksheetCopy()
	{
		_pck.Workbook.Worksheets.Add("Copied Address", _pck.Workbook.Worksheets["Address"]);
		var wsCopy = _pck.Workbook.Worksheets.Add("Copied Comment", _pck.Workbook.Worksheets["Comment"]);

		ExcelPackage pck2 = new();
		pck2.Workbook.Worksheets.Add("Copy From other pck", _pck.Workbook.Worksheets["Address"]);
		pck2.SaveAs(new FileInfo(_worksheetPath + "copy.xlsx"));
		Assert.AreEqual(6, wsCopy.Comments.Count);
	}
	//[Ignore]
	//[TestMethod]
	public void HideTest()
	{
		var ws = _pck.Workbook.Worksheets.Add("Hidden");
		ws.Cells["A1"].Value = "This workbook is hidden";
		ws.Hidden = eWorkSheetHidden.Hidden;
	}
	//[Ignore]
	//[TestMethod]
	public void Hyperlink()
	{
		var ws = _pck.Workbook.Worksheets.Add("HyperLinks");
		var hl = new ExcelHyperLink("G1", "Till G1")
		{
			ToolTip = "Link to cell G1"
		};
		ws.Cells["A1"].Hyperlink = hl;
		//ws.Cells["A2"].Hyperlink = new ExcelHyperLink("mailto:ecsomany@google:huszar", UriKind.Absolute); //Invalid URL will throw an Exception
	}
	//[Ignore]
	//[TestMethod]
	public void VeryHideTest()
	{
		var ws = _pck.Workbook.Worksheets.Add("VeryHidden");
		ws.Cells["a1"].Value = "This workbook is hidden";
		ws.Hidden = eWorkSheetHidden.VeryHidden;
	}
	//[Ignore]
	//[TestMethod]
	public void PrinterSettings()
	{
		var ws = _pck.Workbook.Worksheets.Add("Sod/Hydroseed");

		ws.Cells[1, 1].Value = "1; 1";
		ws.Cells[2, 1].Value = "2; 1";
		ws.Cells[1, 2].Value = "1; 2";
		ws.Cells[2, 2].Value = "2; 2";
		ws.Cells[1, 1, 1, 2].AutoFilter = true;
		ws.PrinterSettings.BlackAndWhite = true;
		ws.PrinterSettings.ShowGridLines = true;
		ws.PrinterSettings.ShowHeaders = true;
		ws.PrinterSettings.PaperSize = ePaperSize.A4;

		ws.PrinterSettings.RepeatRows = new ExcelAddress("1:1");
		ws.PrinterSettings.RepeatColumns = new ExcelAddress("A:A");

		ws.PrinterSettings.Draft = true;
		var r = ws.Cells["A26"];
		r.Value = "X";
		r.Worksheet.Row(26).PageBreak = true;
		ws.PrinterSettings.PrintArea = ws.Cells["A1:B2"];
		ws.PrinterSettings.HorizontalCentered = true;
		ws.PrinterSettings.VerticalCentered = true;

		ws.Select(new ExcelAddress("3:4,E5:F6"));

		ws = _pck.Workbook.Worksheets["RichText"];
		ws.PrinterSettings.RepeatColumns = ws.Cells["A:B"];
		ws.PrinterSettings.RepeatRows = ws.Cells["1:11"];
		ws.PrinterSettings.TopMargin = 1M;
		ws.PrinterSettings.LeftMargin = 1M;
		ws.PrinterSettings.BottomMargin = 1M;
		ws.PrinterSettings.RightMargin = 1M;
		ws.PrinterSettings.Orientation = eOrientation.Landscape;
		ws.PrinterSettings.PaperSize = ePaperSize.A4;
	}
	//[Ignore]
	//[TestMethod]
	public void PictureURL()
	{
		var ws = _pck.Workbook.Worksheets.Add("Pic URL");

		ExcelHyperLink hl = new("http://epplus.codeplex.com")
		{
			ToolTip = "Screen Tip"
		};

		ws.Drawings.AddPicture("Pic URI", EPPlusTest.Properties.Resources.Test1, hl);
	}
	[TestMethod]
	public void TestRepeatRowsAndColumnsTest()
	{
		var p = new ExcelPackage();

		var w = p.Workbook.Worksheets.Add("RepeatRowsAndColumnsTest");

		w.PrinterSettings.RepeatColumns = new ExcelAddress("A:A");
		w.PrinterSettings.RepeatRows = new ExcelAddress("1:1");

		Assert.IsNotNull(w.PrinterSettings.RepeatColumns);
		Assert.IsNotNull(w.PrinterSettings.RepeatRows); // Fails!
	}
	//[Ignore]
	//[TestMethod]
	public void DefaultColWidth()
	{
		var ws = _pck.Workbook.Worksheets.Add("DefColWidth");
		ws.DefaultColWidth = 45;
	}
	[Ignore]
	[TestMethod]
	public void DefColWidthBug()
	{
		var book = _pck.Workbook;
		var sheet = book.Worksheets.Add("Gebruikers");

		sheet.DefaultColWidth = 25d;
		//sheet.defaultRowHeight = 15d; // needed to make sure the resulting file is valid!

		// Create the header row
		sheet.Cells[1, 1].Value = "Afdeling code";
		sheet.Cells[1, 2].Value = "Afdeling naam";
		sheet.Cells[1, 3].Value = "Voornaam";
		sheet.Cells[1, 4].Value = "Tussenvoegsel";
		sheet.Cells[1, 5].Value = "Achternaam";
		sheet.Cells[1, 6].Value = "Gebruikersnaam";
		sheet.Cells[1, 7].Value = "E-mail adres";
		var headerRow = sheet.Cells[1, 1, 1, 7];
		headerRow.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
		headerRow.Style.Font.Size = 12;
		headerRow.Style.Font.Bold = true;
	}
	[Ignore]
	[TestMethod]
	public void ProtectionProblem()
	{
		var xlPackage = new ExcelPackage(new FileInfo("c:\\temp\\CovenantsCheckReportTemplate.xlsx"));
		var ws = xlPackage.Workbook.Worksheets.First();
		ws.Protection.SetPassword("Test");
		xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void Nametest()
	{
		var pck = new ExcelPackage(new FileInfo("c:\\temp\\names.xlsx"));
		var ws = pck.Workbook.Worksheets.First();
		ws.Cells["H37"].Formula = "\"Test\"";
		pck.SaveAs(new FileInfo(@"c:\\temp\\nametest_new.xlsx"));
	}
	//[Ignore]
	[TestMethod]
	public void SetBackground()
	{
		var ws = _pck.Workbook.Worksheets.Add("backimg");

		ws.BackgroundImage.ImageBytes = EPPlusTest.Properties.Resources.Test1;
		ws = _pck.Workbook.Worksheets.Add("backimg2");
		ws.BackgroundImage.SetFromFile(new FileInfo(Path.Combine(_clipartPath, "Vector Drawing.wmf")));
	}
	//[Ignore]
	[TestMethod]
	public void SetHeaderFooterImage()
	{

		var ws = _pck.Workbook.Worksheets.Add("HeaderImage");
		ws.HeaderFooter.OddHeader.CenteredText = "Before ";
		var img = ws.HeaderFooter.OddHeader.InsertPicture(EPPlusTest.Properties.Resources.Test1, PictureAlignment.Centered);
		img.Title = "Renamed Image";

		Assert.AreEqual(426, img.Width);
		img.Width /= 4;
		Assert.AreEqual(49.5, img.Height);
		img.Height /= 4;
		Assert.AreEqual(0, img.Left);
		Assert.AreEqual(0, img.Top);
		ws.HeaderFooter.OddHeader.CenteredText += " After";

		ws.Cells["A1:A400"].Value = 1;

		_pck.Workbook.Worksheets.Copy(ws.Name, "Copied HeaderImage");
	}
	//[Ignore]
	//[TestMethod]
	public void AutoFitColumns()
	{
		var ws = _pck.Workbook.Worksheets.Add("Autofit");
		ws.Cells["A1:H1"].Value = "Auto fit column that is veeery long...";
		ws.Cells["B1"].Style.TextRotation = 30;
		ws.Cells["C1"].Style.TextRotation = 45;
		ws.Cells["D1"].Style.TextRotation = 75;
		ws.Cells["E1"].Style.TextRotation = 90;
		ws.Cells["F1"].Style.TextRotation = 120;
		ws.Cells["G1"].Style.TextRotation = 135;
		ws.Cells["H1"].Style.TextRotation = 180;
		ws.Cells["A1:H1"].AutoFitColumns(0);

		ws.Column(40).AutoFit();
	}
	//[Ignore]
	//[TestMethod]
	public void CopyOverwrite()
	{
		var ws = _pck.Workbook.Worksheets.Add("CopyOverwrite");

		for (var col = 1; col < 15; col++)
		{
			for (var row = 1; row < 30; row++)
			{
				ws.SetValue(row, col, "cell " + ExcelCellBase.GetAddress(row, col));
			}
		}

		ws.Cells["A1:P30"].Copy(ws.Cells["B1"]);
	}
}
