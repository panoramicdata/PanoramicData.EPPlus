using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test;

public partial class CalculationTests
{
	[TestMethod, Ignore]
	public void CalulationValidationExcel()
	{
#if Core
		var dir = AppContext.BaseDirectory;
		dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
		var pck = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "FormulaTest.xlsx")));

		var ws = pck.Workbook.Worksheets["ValidateFormulas"];
		var fr = new Dictionary<string, object>();
		foreach (var cell in ws.Cells)
		{
			if (!string.IsNullOrEmpty(cell.Formula))
			{
				fr.Add(cell.Address, cell.Value);
			}
		}

		pck.Workbook.Calculate();
		var nErrors = 0;
		var errors = new List<Tuple<string, object, object>>();
		foreach (var adr in fr.Keys)
		{
			try
			{
				if (fr[adr] is double && ws.Cells[adr].Value is double)
				{
					var d1 = Convert.ToDouble(fr[adr]);
					var d2 = Convert.ToDouble(ws.Cells[adr].Value);
					if (Math.Abs(d1 - d2) < 0.0001)
					{
						continue;
					}
					else
					{
						Assert.AreEqual(fr[adr], ws.Cells[adr].Value);
					}
				}
				else
				{
					Assert.AreEqual(fr[adr], ws.Cells[adr].Value);
				}
			}
			catch
			{
				errors.Add(new Tuple<string, object, object>(adr, fr[adr], ws.Cells[adr].Value));
				nErrors++;
			}
		}
	}
	[Ignore]
	[TestMethod]
	public void TestOneCell()
	{
		var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Test4.xlsm"));
		_ = pck.Workbook.Worksheets.First();
		pck.Workbook.Worksheets["Räntebärande formaterat utland"].Cells["M13"].Calculate();
		Assert.AreEqual(0d, pck.Workbook.Worksheets["Räntebärande formaterat utland"].Cells["M13"].Value);
	}
	[Ignore]
	[TestMethod]
	public void TestPrecedence()
	{
		var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Precedence.xlsx"));
		var ws = pck.Workbook.Worksheets.Last();
		pck.Workbook.Calculate();
		Assert.AreEqual(150d, ws.Cells["A1"].Value);
	}
	[Ignore]
	[TestMethod]
	public void TestDataType()
	{
		var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusTestark\calc_amount.xlsx"));
		var ws = pck.Workbook.Worksheets.First();
		ws.Names["PRICE"].Value = 30;
		ws.Names["QUANTITY"].Value = 10;

		ws.Calculate();

		ws.Names["PRICE"].Value = 40;
		ws.Names["QUANTITY"].Value = 20;

		ws.Calculate();
	}
}
