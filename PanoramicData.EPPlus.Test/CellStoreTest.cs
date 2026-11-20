using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test;

[TestClass]
public partial class CellStoreTest : TestBase
{

	[TestMethod]
	public void Insert1()
	{
		var ws = _pck.Workbook.Worksheets.Add("Insert1");
		LoadData(ws);

		ws.InsertRow(2, 1000);
		Assert.AreEqual("1,0", ws.GetValue(1002, 1));
		ws.InsertRow(1003, 1000);
		Assert.AreEqual("2,0", ws.GetValue(2003, 1));
		ws.InsertRow(2004, 1000);
		Assert.AreEqual("3,0", ws.GetValue(3004, 1));
		ws.InsertRow(2006, 1000);
		Assert.AreEqual("4,0", ws.GetValue(4005, 1));
		ws.InsertRow(4500, 500);
		Assert.AreEqual("499,0", ws.GetValue(5000, 1));

		ws.InsertRow(1, 1);
		Assert.AreEqual("1,0", ws.GetValue(1003, 1));
		Assert.AreEqual("499,0", ws.GetValue(5001, 1));

		ws.InsertRow(1, 15);
		Assert.AreEqual("3,0", ws.GetValue(4020, 1));
		Assert.AreEqual("499,0", ws.GetValue(5016, 1));

	}
	[TestMethod]
	public void Insert2()
	{
		var ws = _pck.Workbook.Worksheets.Add("Insert2-1");
		LoadData(ws);

		for (var i = 0; i < 32; i++)
		{
			ws.InsertRow(1, 1);
		}

		Assert.AreEqual("0,0", ws.GetValue(33, 1));

		ws = _pck.Workbook.Worksheets.Add("Insert2-2");
		LoadData(ws);

		for (var i = 0; i < 32; i++)
		{
			ws.InsertRow(15, 1);
		}

		Assert.AreEqual("0,0", ws.GetValue(1, 1));
		Assert.AreEqual("14,0", ws.GetValue(47, 1));
	}
	[TestMethod]
	public void Insert3()
	{
		var ws = _pck.Workbook.Worksheets.Add("Insert3");
		LoadData(ws);

		for (var i = 0; i < 500; i += 4)
		{
			ws.InsertRow(i + 1, 2);
		}
	}

	[TestMethod]
	public void InsertRandomTest()
	{
		var ws = _pck.Workbook.Worksheets.Add("Insert4-1");

		LoadData(ws, 5000);

		for (var i = 5000; i > 0; i -= 2)
		{
			ws.InsertRow(i, 1);
		}
	}
	[TestMethod]
	public void EnumCellstore()
	{
		var ws = _pck.Workbook.Worksheets.Add("enum");

		LoadData(ws, 5000);

		var o = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 2, 1, 5, 3);
		foreach (var i in o)
		{
			Console.WriteLine(i);
		}
	}
}
