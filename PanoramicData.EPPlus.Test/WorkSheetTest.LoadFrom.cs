using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Data;
using System.Drawing;
using System.Reflection;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	//[Ignore]
	//[TestMethod]
	public void LoadFromCollectionTest()
	{
		var ws = _pck.Workbook.Worksheets.Add("LoadFromCollection");
		List<TestDTO> list =
		[
			new TestDTO { Id = 1, Name = "Item1", Boolean = false, Date = new DateTime(2011, 1, 1), dto = null, NameVar = "Field 1" },
			new TestDTO { Id = 2, Name = "Item2", Boolean = true, Date = new DateTime(2011, 1, 15), dto = new TestDTO(), NameVar = "Field 2" },
			new TestDTO { Id = 3, Name = "Item3", Boolean = false, Date = new DateTime(2011, 2, 1), dto = null, NameVar = "Field 3" },
		];
		list.Add(new TestDTO { Id = 4, Name = "Item4", Boolean = true, Date = new DateTime(2011, 4, 19), dto = list[1], NameVar = "Field 4" });
		list.Add(new TestDTO { Id = 5, Name = "Item5", Boolean = false, Date = new DateTime(2011, 5, 8), dto = null, NameVar = "Field 5" });
		list.Add(new TestDTO { Id = 6, Name = "Item6", Boolean = true, Date = new DateTime(2010, 3, 27), dto = null, NameVar = "Field 6" });
		list.Add(new TestDTO { Id = 7, Name = "Item7", Boolean = false, Date = new DateTime(2009, 1, 5), dto = list[3], NameVar = "Field 7" });
		list.Add(new TestDTO { Id = 8, Name = "Item8", Boolean = true, Date = new DateTime(2018, 12, 31), dto = null, NameVar = "Field 8" });
		list.Add(new TestDTO { Id = 9, Name = "Item9", Boolean = false, Date = new DateTime(2010, 2, 1), dto = null, NameVar = "Field 9" });

		ws.Cells["A1"].LoadFromCollection(list, true);
		ws.Cells["A30"].LoadFromCollection(list, true, TableStyles.Medium9, BindingFlags.Instance, typeof(TestDTO).GetFields());

		ws.Cells["A45"].LoadFromCollection(list, true, TableStyles.Light1, BindingFlags.Instance, [typeof(TestDTO).GetMethod("GetNameID"), typeof(TestDTO).GetField("NameVar")]);
		ws.Cells["J1"].LoadFromCollection(from l in list where l.Boolean orderby l.Date select new { l.Name, l.Id, l.Date, NameVariable = l.NameVar }, true, TableStyles.Dark4);

		var ints = new[] { 1, 3, 4, 76, 2, 5 };
		ws.Cells["A15"].Value = ints;

		ws = _pck.Workbook.Worksheets.Add("LoadFromCollection_Inherited");
		List<InheritTestDTO> inhList =
		[
			new InheritTestDTO { Id = 1, Name = "Item1", Boolean = false, Date = new DateTime(2011, 1, 1), dto = null, NameVar = "Field 1", InheritedProp = "Inherited 1" },
			new InheritTestDTO { Id = 2, Name = "Item2", Boolean = true, Date = new DateTime(2011, 1, 15), dto = new TestDTO(), NameVar = "Field 2", InheritedProp = "Inherited 2" },
		];
		ws.Cells["A1"].LoadFromCollection(inhList, true);
		Assert.AreEqual("Inherited 2", ws.Cells[3, 1].Value);

		ws.Cells["A5"].LoadFromCollection(inhList, true, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, [typeof(InheritTestDTO).GetProperty("InheritedProp"), typeof(TestDTO).GetProperty("Name")]);
		Assert.AreEqual("Inherited 2", ws.Cells[7, 1].Value);

	}
	//[TestMethod]
	public void LoadFromEmptyCollectionTest()
	{
		_pck ??= new ExcelPackage();
		var ws = _pck.Workbook.Worksheets.Add("LoadFromEmpyCollection");
		List<TestDTO> listDTO = [];
		//List<int> list = new List<int>(0);

		ws.Cells["A1"].LoadFromCollection(listDTO, true);
		ws.Cells["A5"].LoadFromCollection(listDTO, true, TableStyles.Medium9, BindingFlags.Instance, typeof(TestDTO).GetFields());

		ws.Cells["A10"].LoadFromCollection(listDTO, true, TableStyles.Light1, BindingFlags.Instance, [typeof(TestDTO).GetMethod("GetNameID"), typeof(TestDTO).GetField("NameVar")]);
		ws.Cells["A15"].LoadFromCollection(from l in listDTO where l.Boolean orderby l.Date select new { l.Name, l.Id, l.Date, NameVariable = l.NameVar }, true, TableStyles.Dark4);

		ws.Cells["A20"].LoadFromCollection(listDTO, false);
	}
	[TestMethod]
	public void LoadFromOneCollectionTest()
	{
		_pck ??= new ExcelPackage();
		var ws = _pck.Workbook.Worksheets.Add("LoadFromEmpyCollection");
		List<TestDTO> listDTO = new(0) { new TestDTO { Name = "Single" } };

		var r = ws.Cells["A1"].LoadFromCollection(listDTO, true);
		Assert.AreEqual(2, r.Rows);
		var r2 = ws.Cells["A5"].LoadFromCollection(listDTO, false);
		Assert.AreEqual(1, r2.Rows);
	}
	static void Create(string file)
	{
		ExcelPackage pack = new(new FileInfo(file));
		var w = pack.Workbook.Worksheets.Add("delete");
		w.Cells[1, 1].Value = "test";
		w.Cells[1, 2].Value = "test";
		w.Cells[2, 1].Value = "to delete";
		w.Cells[2, 2].Value = "to delete";
		w.Cells[3, 1].Value = "3Left";
		w.Cells[3, 2].Value = "3Left";
		w.Cells[4, 1].Formula = "B3+C3";
		w.Cells[4, 2].Value = "C3+D3";
		pack.Save();
	}

	private static DataTable CreateSampleDataTable(int rowCount)
	{
		object[][] rows = [["Row1", 1, true, 1.5], ["Row2", 2, false, 2.25], ["Row3", 3, true, 3.125]];
		var dt = new DataTable();
		dt.Columns.Add("String", typeof(string));
		dt.Columns.Add("Int", typeof(int));
		dt.Columns.Add("Bool", typeof(bool));
		dt.Columns.Add("Double", typeof(double));
		for (var i = 0; i < rowCount; i++)
		{
			dt.Rows.Add(rows[i]);
		}

		return dt;
	}

	//[TestMethod]
	public void LoadDataReader()
	{
		_pck ??= new ExcelPackage();
		var ws = _pck.Workbook.Worksheets.Add("Loaded DataReader");
		ExcelRangeBase range;
		using var dt = CreateSampleDataTable(2);

		using (var reader = dt.CreateDataReader())
		{
			range = ws.Cells["A1"].LoadFromDataReader(reader, true, "My_Table",
													  TableStyles.Medium5);
		}

		Assert.AreEqual(1, range.Start.Column);
		Assert.AreEqual(4, range.End.Column);
		Assert.AreEqual(1, range.Start.Row);
		Assert.AreEqual(3, range.End.Row);

		using (var reader = dt.CreateDataReader())
		{
			range = ws.Cells["A5"].LoadFromDataReader(reader, false, "My_Table2",
													  TableStyles.Medium5);
		}

		Assert.AreEqual(1, range.Start.Column);
		Assert.AreEqual(4, range.End.Column);
		Assert.AreEqual(5, range.Start.Row);
		Assert.AreEqual(6, range.End.Row);
	}
	//[TestMethod, Ignore]
	public void LoadDataTable()
	{
		_pck ??= new ExcelPackage();
		var ws = _pck.Workbook.Worksheets.Add("Loaded DataTable");

		var dt = CreateSampleDataTable(3);

		ws.Cells["A1"].LoadFromDataTable(dt, true, TableStyles.Medium5);

		ws.Tables[0].Columns[1].TotalsRowFunction = RowFunctions.Sum;
		ws.Tables[0].ShowTotal = true;
	}

	[Ignore]
	[TestMethod]
	public void LoadEmptyDataTable()
	{
		_pck ??= new ExcelPackage();
		var ws = _pck.Workbook.Worksheets.Add("Loaded Empty DataTable");

		var dt = new DataTable();
		dt.Columns.Add(new DataColumn("col1"));
		dt.Columns.Add(new DataColumn("col2"));
		ws.Cells["A1"].LoadFromDataTable(dt, true);

		ws.Cells["D1"].LoadFromDataTable(dt, false);
	}
	[TestMethod]
	public void LoadText_Bug15015()
	{
		var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("Loaded Text");
		ws.Cells["A1"].LoadFromText("\"text with eol,\r\n in a cell\",\"other value\"", new ExcelTextFormat { TextQualifier = '"', EOL = ",\r\n", Delimiter = ',' });
	}

	[TestMethod]
	public void LoadText_Bug15015_Negative()
	{
		var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("Loaded Text");
		var exceptionThrown = false;
		try
		{
			ws.Cells["A1"].LoadFromText("\"text with eol,\r\n",
										new ExcelTextFormat { TextQualifier = '"', EOL = ",\r\n", Delimiter = ',' });
		}
		catch
		{
			exceptionThrown = true;
		}

		Assert.IsTrue(exceptionThrown, "Exception thrown");
	}

	//[Ignore]
	[TestMethod]
	public void LoadText()
	{
		var ws = _pck.Workbook.Worksheets.Add("Loaded Text");

		ws.Cells["A1"].LoadFromText("1.2");
		ws.Cells["A2"].LoadFromText("1,\"Test av data\",\"12,2\",\"\"Test\"\"");
		ws.Cells["A3"].LoadFromText("\"1,3\",\"Test av \"\"data\",\"12,2\",\"Test\"\"\"", new ExcelTextFormat { TextQualifier = '"' });
		ws.Cells["A4"].LoadFromText("\"1,3\",\"\",\"12,2\",\"Test\"\"\"", new ExcelTextFormat { TextQualifier = '"' });

		ws = _pck.Workbook.Worksheets.Add("File1");
		if (File.Exists(@"c:\temp\csv\et1c1004.csv"))
		{
			ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\et1c1004.csv"), new ExcelTextFormat { SkipLinesBeginning = 3, SkipLinesEnd = 1, EOL = "\n" });
		}

		ws = _pck.Workbook.Worksheets.Add("File2");
		if (File.Exists(@"c:\temp\csv\etiv2812.csv"))
		{
			ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\etiv2812.csv"), new ExcelTextFormat { SkipLinesBeginning = 3, SkipLinesEnd = 1, EOL = "\n" });
		}

		ws = _pck.Workbook.Worksheets.Add("File3");
		if (File.Exists(@"c:\temp\csv\last_gics.txt"))
		{
			ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\last_gics.txt"), new ExcelTextFormat { SkipLinesBeginning = 1, Delimiter = '|' });
		}

		ws = _pck.Workbook.Worksheets.Add("File4");

		ws.Cells["A1"].LoadFromText("1,\"Test\",\"\",\"\"\"\",3\r\n", new ExcelTextFormat { TextQualifier = '\"' });

		var style = _pck.Workbook.Styles.CreateNamedStyle("RedStyle");
		style.Style.Fill.PatternType = ExcelFillStyle.Solid;
		style.Style.Fill.BackgroundColor.SetColor(Color.Red);
	}
}
