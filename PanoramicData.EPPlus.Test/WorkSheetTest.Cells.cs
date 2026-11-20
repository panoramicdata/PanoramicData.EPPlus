using OfficeOpenXml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	//[Ignore]
	//[TestMethod]
	public void RichTextCells()
	{
		var ws = _pck.Workbook.Worksheets.Add("RichText");
		var rs = ws.Cells["A1"].RichText;

		var r1 = rs.Add("Test");
		r1.Bold = true;
		r1.Color = Color.Pink;

		var r2 = rs.Add(" of");
		r2.Size = 14;
		r2.Italic = true;

		var r3 = rs.Add(" rich");
		r3.FontName = "Arial";
		r3.Size = 18;
		r3.Italic = true;

		var r4 = rs.Add("text.");
		r4.Size = 8.25f;
		r4.Italic = true;
		r4.UnderLine = true;

		var rIns = rs.Insert(2, " inserted");
		rIns.Bold = true;
		rIns.Color = Color.Green;

		rs = ws.Cells["A3:A4"].RichText;

		var r5 = rs.Add("Double");
		r5.Color = Color.PeachPuff;
		r5.FontName = "times new roman";
		r5.Size = 16;

		var r6 = rs.Add(" cells");
		r6.Color = Color.Red;
		r6.UnderLine = true;


		rs = ws.Cells["C8"].RichText;
		r1 = rs.Add("Blue ");
		r1.Color = Color.Blue;

		r2 = rs.Add("Red");
		r2.Color = Color.Red;

		ws.Cells["G1"].RichText.Add("Room 02 & 03");
		ws.Cells["G2"].RichText.Text = "Room 02 & 03";

		ws = _pck.Workbook.Worksheets.Add("RichText2");
		ws.Cells["A1"].RichText.Text = "Room 02 & 03";
		ws.TabColor = Color.PowderBlue;

		r1 = ws.Cells["G3"].RichText.Add("Test");
		r1.Bold = true;
		ws.Cells["G3"].RichText.Add(" a new t");
		ws.Cells["G3"].RichText[1].Bold = false;
	}
	//[Ignore]
	//[TestMethod]
	public void TestComments()
	{
		var ws = _pck.Workbook.Worksheets.Add("Comment");
		var comment = ws.Comments.Add(ws.Cells["C3"], "Jan Källman\r\nAuthor\r\n", "JK");
		comment.RichText[0].Bold = true;
		comment.RichText[0].PreserveSpace = true;
		comment.RichText.Add("Test comment");
		comment.VerticalAlignment = eTextAlignVerticalVml.Center;
		ws.Comments.Add(ws.Cells["A2"], "Jan Källman\r\nAuthor\r\n1", "JK");

		comment = ws.Comments.Add(ws.Cells["A1"], "Jan Källman\r\nAuthor\r\n2", "JK");
		comment.AlternativeText = "Test of AlternetiveText2";
		ws.Comments.Add(ws.Cells["C2"], "Jan Källman\r\nAuthor\r\n3", "JK");
		ws.Comments.Add(ws.Cells["C1"], "Jan Källman\r\nAuthor\r\n5", "JK");
		ws.Comments.Add(ws.Cells["B1"], "Jan Källman\r\nAuthor\r\n7", "JK");

		ws.Comments.Remove(ws.Cells["A2"].Comment);

		ws.Cells["B2"].AddComment("Range Added Comment test test test test test test test test test test testtesttesttesttesttesttesttesttesttesttest", "Jan Källman");
		ws.Cells["c3"].Comment.AutoFit = true;

	}
	//[Ignore]
	//[TestMethod]
	public void Address()
	{
		var ws = _pck.Workbook.Worksheets.Add("Address");
		ws.Cells["A1:A4,B5:B7"].Value = "AddressTest";
		ws.Cells["A1:A4,B5:B7"].Style.Font.Color.SetColor(Color.Red);
		ws.Cells["A2:A3,B4:B8"].Style.Fill.PatternType = ExcelFillStyle.LightUp;
		ws.Cells["A2:A3,B4:B8"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
		ws.Cells["2:2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
		ws.Cells["2:2"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
		ws.Cells["B:B"].Style.Font.Name = "Times New Roman";

		ws.Cells["C4:G4,H8:H30,B15"].FormulaR1C1 = "RC[-1]+R1C[-1]";
		ws.Cells["C4:G4,H8:H30,B15"].Style.Numberformat.Format = "#,##0.000";
		ws.Cells["G1,G3"].Hyperlink = new ExcelHyperLink("Comment!$A$1", "Comment");
		ws.Cells["G1,G3"].Style.Font.Color.SetColor(Color.Blue);
		ws.Cells["G1,G3"].Style.Font.UnderLine = true;

		ws.Cells["A1:G5"].Copy(ws.Cells["A50"]);

		var ws2 = _pck.Workbook.Worksheets.Add("Copy Cells");
		ws.Cells["1:4"].Copy(ws2.Cells["1:1"]);

		ws.Cells["H1:J5"].Merge = true;
		ws.Cells["2:3"].Copy(ws.Cells["50:51"]);

		ws.Cells["A40"].Value = new string([(char)8, (char)9]);

		var styleRng = ws.Cells["A1"];
		var tempStyle = styleRng.Style;
		var namedStyle = _pck.Workbook.Styles.CreateNamedStyle("HyperLink", tempStyle);
		namedStyle.Style.Font.UnderLineType = ExcelUnderLineType.Single;
		namedStyle.Style.Font.Color.SetColor(Color.Blue);
	}
	//[Ignore]
	//[TestMethod]
	public void Encoding()
	{
		var ws = _pck.Workbook.Worksheets.Add("Encoding");
		ws.Cells["A1"].Value = "_x0099_";
		ws.Cells["A2"].Value = " Test \b" + (char)1 + " end\"";
		ws.Cells["A3"].Value = "_x0097_ test_x001D_1234";
		ws.Cells["A4"].Value = "test" + (char)31;   //Bug issue 14689 //Fixed
	}
	//[Ignore]
	//[TestMethod]
	public void ValueError()
	{
		var ws = _pck.Workbook.Worksheets.Add("ValueError");

		ws.Cells[1, 1].Value = "Domestic Violence&#xB; and the Professional";
		var rt = ws.Cells[1, 2].RichText.Add("Domestic Violence&#xB; and the Professional 2");
		Console.WriteLine(rt.Bold.ToString());
		rt.Bold = true;
		Console.WriteLine(rt.Bold.ToString());
	}
	//[Ignore]
	//[TestMethod]
	public void FormulaError()
	{
		var ws = _pck.Workbook.Worksheets.Add("FormulaError");

		ws.Cells["D5"].Formula = "COUNTIF(A1:A100,\"Miss\")";
		ws.Cells["A1:K3"].Formula = "A3+A4";
		ws.Cells["A4"].FormulaR1C1 = "+ROUNDUP(RC[1]/10,0)*10";

		ws = _pck.Workbook.Worksheets.Add("Sheet-RC1");
		ws.Cells["A4"].FormulaR1C1 = "+ROUNDUP('Sheet-RC1'!RC[1]/10,0)*10";

		//ws.Cells["B2:I2"].Formula = "";   //Error
	}
	[TestMethod, Ignore]
	public void FormulaArray()
	{
		_pck = new ExcelPackage();
		var ws = _pck.Workbook.Worksheets.Add("FormulaError");

		ws.Cells["E2:E5"].CreateArrayFormula("FREQUENCY(B2:B18,C2:C5)");
		_pck.SaveAs(new FileInfo("c:\\temp\\arrayformula.xlsx"));
	}
	//[Ignore]
	//[TestMethod]
	public void FormulaOverwrite()
	{
		var ws = _pck.Workbook.Worksheets.Add("FormulaOverwrite");
		//Inside
		ws.Cells["A1:G12"].Formula = "B1+C1";
		ws.Cells["B2:C3"].Formula = "G2+E1";


		//Top bottom overwrite
		ws.Cells["A14:G26"].Formula = "B1+C1+D1";
		ws.Cells["B13:C28"].Formula = "G2+E1";

		//Top bottom overwrite
		ws.Cells["B30:E42"].Formula = "B1+C1+$D$1";
		ws.Cells["A32:H33"].Formula = "G2+E1";

		ws.Cells["A50:A59"].CreateArrayFormula("C50+D50");
		ws.Cells["A1"].Value = "test";
		ws.Cells["A15"].Value = "Värde";
		ws.Cells["C12"].AddComment("Test", "JJOD");
		ws.Cells["D12:I12"].Merge = true;
		ws.Cells["D12:I12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
		ws.Cells["D12:I12"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
		ws.Cells["D12:I12"].Style.WrapText = true;

		ws.Cells["F1:F3"].Formula = "F2+F3";
		ws.Cells["J1:J3"].Formula = "F2+F3";
		ws.Cells["F1:F3"].Formula = "F5+F6";    //Overwrite same range
	}
	//[Ignore]
	//[TestMethod]
	public void DefinedName()
	{
		var ws = _pck.Workbook.Worksheets.Add("Names");
		ws.Names.Add("RefError", ws.Cells["#REF!"]);

		ws.Cells["A1"].Value = "Test";
		ws.Cells["A1"].Style.Font.Size = 8.5F;

		ws.Names.Add("Address", ws.Cells["A2:A3"]);
		ws.Cells["Address"].Value = 1;
		ws.Names.AddValue("Value", 5);
		ws.Names.Add("FullRow", ws.Cells["2:2"]);
		ws.Names.Add("FullCol", ws.Cells["A:A"]);

		ws.Names.AddFormula("Formula", "Names!A2+Names!A3+Names!Value");
	}
	[Ignore]
	[TestMethod]
	public void URL()
	{
		var p = new ExcelPackage(new FileInfo(@"c:\temp\url.xlsx"));
		foreach (var ws in p.Workbook.Worksheets)
		{
			// Enumerating must not throw.
		}

		p.SaveAs(new FileInfo(@"c:\temp\urlsaved.xlsx"));
	}
	//[Ignore]
	//[TestMethod]
	public void Merge()
	{
		var ws = _pck.Workbook.Worksheets.Add("Merge");
		ws.Cells["A1:A4"].Merge = true;
		ws.Cells["C1:C4,C8:C12"].Merge = true;
		ws.Cells["D13:E18,G5,U32:U45"].Merge = true;
		ws.Cells["D13:E18,G5,U32:U45"].Style.WrapText = true;
		ws.Cells["AA:AC"].Merge = true;
		ws.SetValue(13, 4, "Merged\r\nnew row");
	}
	//[Ignore]
	//[TestMethod]
	public void LoadArray()
	{
		var ws = _pck.Workbook.Worksheets.Add("Loaded Array");
		List<object[]> testArray = [[3, 4, 5, 6], ["Test1", "test", "5", "6"]];
		ws.Cells["A1"].LoadFromArrays(testArray);
	}
	[Ignore]
	[TestMethod]
	public void Mergebug()
	{
		var xlPackage = new ExcelPackage();
		var xlWorkSheet = xlPackage.Workbook.Worksheets.Add("Test Sheet");
		var Cells = xlWorkSheet.Cells;
		var TitleCell = Cells[1, 1, 1, 3];

		TitleCell.Merge = true;
		TitleCell.Value = "Test Spreadsheet";
		Cells[2, 1].Value = "Test Sub Heading\r\ntest" + (char)22;
		for (var i = 0; i < 256; i++)
		{
			Cells[3, i + 1].Value = (char)i;
		}

		Cells[2, 1].Style.WrapText = true;
		xlWorkSheet.Row(1).Height = 50;
		xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
	}
}
