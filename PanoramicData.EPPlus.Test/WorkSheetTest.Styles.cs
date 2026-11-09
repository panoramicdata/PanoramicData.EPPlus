using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	[Ignore]
	[TestMethod]
	public void RowStyle()
	{
		FileInfo newFile = new(_worksheetPath + @"sample8.xlsx");
		if (newFile.Exists)
		{
			newFile.Delete();  // ensures we create a new workbook
		}

		ExcelPackage package = new();
		//Load the sheet with one string column, one date column and a few random numbers.
		var ws = package.Workbook.Worksheets.Add("First line test");

		ws.Cells[1, 1].Value = "1; 1";
		ws.Cells[2, 1].Value = "2; 1";
		ws.Cells[1, 2].Value = "1; 2";
		ws.Cells[2, 2].Value = "2; 2";

		ws.Row(1).Style.Font.Bold = true;
		ws.Column(1).Style.Font.Bold = true;
		package.SaveAs(newFile);

	}
	//[Ignore]
	//[TestMethod]
	public void StyleNameTest()
	{
		var ws = _pck.Workbook.Worksheets.Add("StyleNameTest");

		ws.Cells[1, 1].Value = "R1 C1";
		ws.Cells[1, 2].Value = "R1 C2";
		ws.Cells[1, 3].Value = "R1 C3";
		ws.Cells[2, 1].Value = "R2 C1";
		ws.Cells[2, 2].Value = "R2 C2";
		ws.Cells[2, 3].Value = "R2 C3";
		ws.Cells[3, 1].Value = double.PositiveInfinity;
		ws.Cells[3, 2].Value = double.NegativeInfinity;
		ws.Cells[4, 1].CreateArrayFormula("A1+B1");
		var ns = _pck.Workbook.Styles.CreateNamedStyle("TestStyle");
		ns.Style.Font.Bold = true;

		ws.Cells.Style.Locked = true;
		ws.Cells["A1:C1"].StyleName = "TestStyle";
		ws.DefaultRowHeight = 35;
		ws.Cells["A1:C4"].Style.Locked = false;
		ws.Protection.IsProtected = true;
	}
	[Ignore]
	[TestMethod]
	public void Stylebug()
	{
		ExcelPackage p = new(new FileInfo(@"c:\temp\FullProjecte.xlsx"));

		var ws = p.Workbook.Worksheets.First();
		ws.Cells[12, 1].Value = 0;
		ws.Cells[12, 2].Value = new DateTime(2010, 9, 14);
		ws.Cells[12, 3].Value = "Federico Lois";
		ws.Cells[12, 4].Value = "Nakami";
		ws.Cells[12, 5].Value = "Hores";
		ws.Cells[12, 7].Value = 120;
		ws.Cells[12, 8].Value = "A definir";
		ws.Cells[12, 9].Value = new DateTime(2010, 9, 14);
		ws.Cells[12, 10].Value = new DateTime(2010, 9, 14);
		ws.Cells[12, 11].Value = "Transferència";

		ws.InsertRow(13, 1, 12);
		ws.Cells[13, 1].Value = 0;
		ws.Cells[13, 2].Value = new DateTime(2010, 9, 14);
		ws.Cells[13, 3].Value = "Federico Lois";
		ws.Cells[13, 4].Value = "Nakami";
		ws.Cells[13, 5].Value = "Hores";
		ws.Cells[13, 7].Value = 120;
		ws.Cells[13, 8].Value = "A definir";
		ws.Cells[13, 9].Value = new DateTime(2010, 9, 14);
		ws.Cells[13, 10].Value = new DateTime(2010, 9, 14);
		ws.Cells[13, 11].Value = "Transferència";

		ws.InsertRow(14, 1, 13);

		ws.InsertRow(19, 1, 19);
		ws.InsertRow(26, 1, 26);
		ws.InsertRow(33, 1, 33);
		p.SaveAs(new FileInfo(@"c:\temp\FullProjecte_new.xlsx"));
	}
	//[Ignore]
	//[TestMethod]
	public void NamedStyles()
	{
		var wsSheet = _pck.Workbook.Worksheets.Add("NamedStyles");

		var firstNamedStyle =
			_pck.Workbook.Styles.CreateNamedStyle("templateFirst");

		var s = firstNamedStyle.Style;

		s.Fill.PatternType = ExcelFillStyle.Solid;
		s.Fill.BackgroundColor.SetColor(Color.LightGreen);
		s.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
		s.VerticalAlignment = ExcelVerticalAlignment.Center;

		var secondNamedStyle = _pck.Workbook.Styles.CreateNamedStyle("first", firstNamedStyle.Style).Style;
		secondNamedStyle.Font.Bold = true;
		secondNamedStyle.Font.Name = "Arial Black";
		secondNamedStyle.Font.Size = 8;
		secondNamedStyle.Border.Bottom.Style = ExcelBorderStyle.Medium;
		secondNamedStyle.Border.Left.Style = ExcelBorderStyle.Medium;

		wsSheet.Cells["B2"].Value = "Text Center";
		wsSheet.Cells["B2"].StyleName = "first";
		_pck.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";

		var rowStyle = _pck.Workbook.Styles.CreateNamedStyle("RowStyle", firstNamedStyle.Style).Style;
		rowStyle.Fill.BackgroundColor.SetColor(Color.Pink);
		wsSheet.Cells.StyleName = "templateFirst";
		wsSheet.Cells["C5:H15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
		wsSheet.Cells["C5:H15"].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);

		wsSheet.Cells["30:35"].StyleName = "RowStyle";
		var colStyle = _pck.Workbook.Styles.CreateNamedStyle("columnStyle", firstNamedStyle.Style).Style;
		colStyle.Fill.BackgroundColor.SetColor(Color.CadetBlue);

		wsSheet.Cells["D:E"].StyleName = "ColumnStyle";

		//Verify that cloned styles works.
		var defaultStyle = _pck.Workbook.Styles.CreateNamedStyle("Default");
		defaultStyle.Style.Font.Name = "Arial";
		defaultStyle.Style.Font.Size = 18;
		defaultStyle.Style.Font.UnderLine = true;
		var boldStyle = _pck.Workbook.Styles.CreateNamedStyle("Bold", defaultStyle.Style);
		boldStyle.Style.Font.Color.SetColor(Color.Red);

		Assert.AreEqual("Arial", defaultStyle.Style.Font.Name);
		Assert.AreEqual(18, defaultStyle.Style.Font.Size);

		Assert.AreEqual("FFFF0000", boldStyle.Style.Font.Color.Rgb);

	}
	//[Ignore]
	//[TestMethod]
	public void StyleFill()
	{
		var ws = _pck.Workbook.Worksheets.Add("Fills");
		ws.Cells["A1:C3"].Style.Fill.Gradient.Type = ExcelFillGradientType.Linear;
		ws.Cells["A1:C3"].Style.Fill.Gradient.Color1.SetColor(Color.Red);
		ws.Cells["A1:C3"].Style.Fill.Gradient.Color2.SetColor(Color.Blue);

		ws.Cells["J20:J23"].Style.Fill.PatternType = ExcelFillStyle.Solid;
		ws.Cells["J20:J23"].Style.Fill.BackgroundColor.SetColor(0xFF, 0x00, 0XFF, 0x00); //Green

		ws.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.MediumGray;
		ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
		var r = ws.Cells["A2:A3"];
		r.Style.Fill.Gradient.Type = ExcelFillGradientType.Path;
		r.Style.Fill.Gradient.Left = 0.7;
		r.Style.Fill.Gradient.Right = 0.7;
		r.Style.Fill.Gradient.Top = 0.7;
		r.Style.Fill.Gradient.Bottom = 0.7;

		ws.Cells[4, 1, 4, 360].Style.Fill.Gradient.Type = ExcelFillGradientType.Path;

		for (double col = 1; col < 360; col++)
		{
			r = ws.Cells[4, Convert.ToInt32(col)];
			r.Style.Fill.Gradient.Degree = col;
			r.Style.Fill.Gradient.Left = col / 360;
			r.Style.Fill.Gradient.Right = col / 360;
			r.Style.Fill.Gradient.Top = col / 360;
			r.Style.Fill.Gradient.Bottom = col / 360;
		}

		r = ws.Cells["A5"];
		r.Style.Fill.Gradient.Left = .50;

		ws = _pck.Workbook.Worksheets.Add("FullFills");
		ws.Cells.Style.Fill.Gradient.Left = 0.25;
		ws.Cells["A1"].Value = "test";
		ws.Cells["A1"].RichText.Add("Test rt");
		ws.Cells.AutoFilter = true;
		Assert.AreNotEqual(null, ws.Cells["A1:D5"].Value);
	}
	[Ignore]
	[TestMethod]
	public void BuildInStyles()
	{
		var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("Default");
		ws.Cells.Style.Font.Name = "Arial";
		ws.Cells.Style.Font.Size = 15;
		ws.Cells.Style.Border.Bottom.Style = ExcelBorderStyle.MediumDashed;
		var n = pck.Workbook.Styles.NamedStyles[0];
		n.Style.Numberformat.Format = "yyyy";
		n.Style.Font.Name = "Arial";
		n.Style.Font.Size = 15;
		n.Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;
		n.Style.Border.Bottom.Color.SetColor(Color.Red);
		n.Style.Fill.PatternType = ExcelFillStyle.Solid;
		n.Style.Fill.BackgroundColor.SetColor(Color.Blue);
		n.Style.Border.Bottom.Color.SetColor(Color.Red);
		n.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
		n.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
		n.Style.TextRotation = 90;
		ws.Cells["a1:c3"].StyleName = "Normal";
		pck.SaveAs(new FileInfo(@"c:\temp\style.xlsx"));
	}
}
