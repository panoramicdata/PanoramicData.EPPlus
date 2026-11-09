using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[TestMethod]
	public void Issue170()
	{
		OpenTemplatePackage("print_titles_170.xlsx");
		_pck.Compatibility.IsWorksheets1Based = false;
		var sheet = _pck.Workbook.Worksheets[0];

		sheet.PrinterSettings.RepeatColumns = new ExcelAddress("$A:$C");
		sheet.PrinterSettings.RepeatRows = new ExcelAddress("$1:$3");

		SaveWorksheet("print_titles_170-Saved.xlsx");
		_pck.Dispose();
	}
	[TestMethod]
	public void Issue172()
	{
		OpenTemplatePackage("quest.xlsx");
		foreach (var ws in _pck.Workbook.Worksheets)
		{
			Console.WriteLine(ws.Name);
		}

		_pck.Dispose();
	}

	[TestMethod]
	public void Issue219()
	{
		OpenTemplatePackage("issueFile.xlsx");
		foreach (var ws in _pck.Workbook.Worksheets)
		{
			Console.WriteLine(ws.Name);
		}

		_pck.Dispose();
	}
	[TestMethod]
	public void Issue234()
	{
		using var s = new MemoryStream();
		var data = Encoding.UTF8.GetBytes("Bad data").ToArray();
		s.Write(data, 0, data.Length);
		Assert.ThrowsExactly<InvalidDataException>(() => new ExcelPackage(s));
	}

	[TestMethod]
	public void Issue220()
	{
		OpenPackage("sheetname_pbl.xlsx", true);
		var ws = _pck.Workbook.Worksheets.Add("Deal's History");
		var a = ws.Cells["A:B"];
		ws.AutoFilterAddress = ws.Cells["A1:C3"];
		_pck.Workbook.Names.Add("Test", ws.Cells["B1:D2"]);
		_ = a.WorkSheet;

		var a2 = new ExcelAddress("'Deal''s History'!a1:a3");
		Assert.AreEqual("Deal's History", a2.WorkSheet);
		_pck.Save();
		_pck.Dispose();

	}
	[TestMethod]
	public void Issue233()
	{
		//get some test data
		var cars = Car.GenerateList();

		OpenPackage("issue233.xlsx", true);

		var sheetName = "Summary_GLEDHOWSUGARCO![]()PTY";

		//Create the worksheet 
		var sheet = _pck.Workbook.Worksheets.Add(sheetName);

		//Read the data into a range
		var range = sheet.Cells["A1"].LoadFromCollection(cars, true);

		//Make the range a table: the sheet name's characters make the table name invalid
		Assert.ThrowsExactly<ArgumentException>(() => sheet.Tables.Add(range, $"data{sheetName}"));
	}
	[TestMethod]
	public void Issue236()
	{
		OpenTemplatePackage("Issue236.xlsx");
		_pck.Workbook.Worksheets["Sheet1"].Cells[7, 10].AddComment("test", "Author");
		SaveWorksheet("Issue236-Saved.xlsx");
	}
	[TestMethod]
	public void Issue228()
	{
		OpenTemplatePackage("Font55.xlsx");
		var ws = _pck.Workbook.Worksheets["Sheet1"];
		ws.Drawings.AddShape("Shape1", eShapeStyle.Diamond);
		ws.Cells["A1"].Value = "tasetraser";
		ws.Cells.AutoFitColumns();
		SaveWorksheet("Font55-Saved.xlsx");
	}
	[TestMethod]
	public void Issue241()
	{
		OpenPackage("issue241", true);
		var wks = _pck.Workbook.Worksheets.Add("test");
		wks.DefaultRowHeight = 35;
		_pck.Save();
		_pck.Dispose();
	}
	[TestMethod]
	public void Issue195()
	{
		var pkg = new ExcelPackage();
		pkg.Workbook.Worksheets.Add("Sheet1");
		var defaultStyle = pkg.Workbook.Styles.CreateNamedStyle("Default");
		defaultStyle.Style.Font.Name = "Arial";
		defaultStyle.Style.Font.Size = 18;
		defaultStyle.Style.Font.UnderLine = true;
		var boldStyle = pkg.Workbook.Styles.CreateNamedStyle("Bold", defaultStyle.Style);
		boldStyle.Style.Font.Color.SetColor(Color.Red);

		Assert.AreEqual("Arial", defaultStyle.Style.Font.Name);
		Assert.AreEqual(18, defaultStyle.Style.Font.Size);

		Assert.AreEqual("Arial", boldStyle.Style.Font.Name);
		Assert.AreEqual(18, boldStyle.Style.Font.Size);
		Assert.AreEqual("FFFF0000", boldStyle.Style.Font.Color.Rgb);

		pkg.SaveAs(new FileInfo(@"c:\temp\n.xlsx"));
	}
	[TestMethod]
	public void Issue332()
	{
		InitBase();
		var pkg = OpenPackage("Hyperlink.xlsx", true);
		var ws = pkg.Workbook.Worksheets.Add("Hyperlink");
		ws.Cells["A1"].Hyperlink = new ExcelHyperLink("A2", "A2");
		pkg.Save();
	}
	[TestMethod]
	public void Issue332_2()
	{
		InitBase();
		var pkg = OpenPackage("Hyperlink.xlsx");
		var ws = pkg.Workbook.Worksheets["Hyperlink"];
		Assert.IsNotNull(ws.Cells["A1"].Hyperlink);
	}
	[TestMethod]
	public void Issuer246()
	{
		InitBase();
		OpenPackage("issue246.xlsx", true);
		var ws = _pck.Workbook.Worksheets.Add("DateFormat");
		ws.Cells["A1"].Value = 43465;
		ws.Cells["A1"].Style.Numberformat.Format = @"[$-F800]dddd,\ mmmm\ dd,\ yyyy";
		_pck.Save();

		OpenPackage("issue246.xlsx");
		ws = _pck.Workbook.Worksheets["DateFormat"];
		var pCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
		System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-SE", false);
		var sv = new CultureInfo("sv-SE", false);
		Assert.AreEqual(new DateTime(2018, 12, 31).ToString("D", sv), ws.Cells["A1"].Text);
		Assert.AreEqual(ws.GetValue<DateTime>(1, 1), new DateTime(2018, 12, 31));
		System.Threading.Thread.CurrentThread.CurrentCulture = pCulture;
	}
	[TestMethod]
	public void Issue347()
	{
		var package = OpenTemplatePackage("Issue327.xlsx");
		var templateWS = package.Workbook.Worksheets["Template"];
		package.Workbook.Worksheets.Delete(templateWS);
	}
	[TestMethod]
	public void Issue348()
	{
		using ExcelPackage pck = new();
		var ws = pck.Workbook.Worksheets.Add("S1");
		var formula = "VLOOKUP(C2,A:B,1,0)";
		ws.Cells[2, 4].Formula = formula;
		_ = ws.Cells[2, 4].FormulaR1C1; // VLOOKUP(C2,C[-3]:C[-2],1,0)
		ws.Cells[2, 5].FormulaR1C1 = ws.Cells[2, 4].FormulaR1C1;
		_ = ws.Cells[2, 5].FormulaR1C1; // VLOOKUP(C2,C[-3]**:B:C:C**,1,0)   //unexpected value here
	}

	[TestMethod, Ignore]
	public void Issue376()
	{
		using ExcelPackage pck = new();
		var ws = pck.Workbook.Worksheets.Add("test");
		var v = ws.DataValidations.AddListValidation("A1");
		v.Formula.Values.Add("1");
		v.Formula.Values.Add("2");
		v.ShowErrorMessage = true;
		v.Error = "error!";
		v.ErrorStyle = OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle.stop;
		v.AllowBlank = false;
		pck.SaveAs(new FileInfo(@"c:\temp\book.xlsx"));
	}

	[TestMethod]
	public void Issue367()
	{
		using var pck = OpenTemplatePackage(@"ProductFunctionTest.xlsx");
		var sheet = pck.Workbook.Worksheets.First();
		sheet.Cells["B14"].Value = 11;
		sheet.Cells["B15"].Value = 13;
		sheet.Cells["B16"].Formula = "Product(B13:B15)";
		sheet.Calculate();

		Assert.AreEqual(0d, sheet.Cells["B16"].Value);
	}
}
