using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System.Text;

namespace PanoramicData.EPPlus.Test;

public partial class CalculationTests
{
	public static void TestAllWorkbooks()
	{
		StringBuilder sb = new();
		//Add sheets to test in this directory or change it to your testpath.
		var path = @"C:\temp\EPPlusTestark\workbooks";
		if (!Directory.Exists(path))
		{
			return;
		}

		foreach (var file in Directory.GetFiles(path, "*.xls*"))
		{
			sb.Append(GetOutput(file));
		}

		if (sb.Length > 0)
		{
			File.WriteAllText(string.Format("TestAllWorkooks{0}.txt", DateTime.Now.ToString("d") + " " + DateTime.Now.ToString("t")), sb.ToString());
			throw new InvalidOperationException("Test failed with\r\n\r\n" + sb);
		}
	}
	[TestMethod]
	public void CalculateDateMath()
	{
		using ExcelPackage package = new();
		var worksheet = package.Workbook.Worksheets.Add("Test");
		var dateCell = worksheet.Cells[2, 2];
		var date = new DateTime(2013, 1, 1);
		dateCell.Value = date;
		var quotedDateCell = worksheet.Cells[2, 3];
		quotedDateCell.Formula = $"\"{date:d}\"";
		var dateFormula = "B2";
		var dateFormulaWithMath = "B2+1";
		var quotedDateFormulaWithMath = $"\"{date:d}\"+1";
		var quotedDateReferenceFormulaWithMath = "C2+1";
		var expectedDate = 41275.0; // January 1, 2013
		var expectedDateWithMath = 41276.0; // January 2, 2013
		Assert.AreEqual(expectedDate, worksheet.Calculate(dateFormula));
		Assert.AreEqual(expectedDateWithMath, worksheet.Calculate(dateFormulaWithMath));
		Assert.AreEqual(expectedDateWithMath, worksheet.Calculate(quotedDateFormulaWithMath));
		Assert.AreEqual(expectedDateWithMath, worksheet.Calculate(quotedDateReferenceFormulaWithMath));
		var formulaCell = worksheet.Cells[2, 4];
		formulaCell.Formula = dateFormulaWithMath;
		formulaCell.Calculate();
		Assert.AreEqual(expectedDateWithMath, formulaCell.Value);
		formulaCell.Formula = quotedDateReferenceFormulaWithMath;
		formulaCell.Calculate();
		Assert.AreEqual(expectedDateWithMath, formulaCell.Value);
	}
	private static string GetOutput(string file)
	{
		using var pck = new ExcelPackage(new FileInfo(file));
		var fr = new Dictionary<string, object>();
		foreach (var ws in pck.Workbook.Worksheets)
		{
			if (ws is not ExcelChartsheet)
			{
				foreach (var cell in ws.Cells)
				{
					if (!string.IsNullOrEmpty(cell.Formula))
					{
						fr.Add(ws.PositionID.ToString() + "," + cell.Address, cell.Value);
						ws.SetValueInner(cell.Start.Row, cell.Start.Column, null);
					}
				}
			}
		}

		pck.Workbook.Calculate();
		var nErrors = 0;
		ExcelWorksheet sheet = null;
		var adr = "";
		var fileErr = new StreamWriter(new FileStream("c:\\temp\\err.txt", FileMode.Append));
		foreach (var cell in fr.Keys)
		{
			try
			{
				var spl = cell.Split(',');
				var ix = int.Parse(spl[0]);
				sheet = pck.Workbook.Worksheets[ix];
				adr = spl[1];
				if (fr[cell] is double && (sheet.Cells[adr].Value is double || sheet.Cells[adr].Value is decimal || OfficeOpenXml.Compatibility.TypeCompat.IsPrimitive(sheet.Cells[adr].Value)))
				{
					var d1 = Convert.ToDouble(fr[cell]);
					var d2 = Convert.ToDouble(sheet.Cells[adr].Value);
					if (Equals(d1, d2))
					{
						continue;
					}
					else
					{
						fileErr.WriteLine("Diff cell " + sheet.Name + "!" + adr + "\t" + d1.ToString("R15") + "\t" + d2.ToString("R15"));
					}
				}
				else
				{
					if ((fr[cell] ?? "").ToString() != (sheet.Cells[adr].Value ?? "").ToString())
					{
						fileErr.WriteLine("String?  cell " + sheet.Name + "!" + adr + "\t" + (fr[cell] ?? "").ToString() + "\t" + (sheet.Cells[adr].Value ?? "").ToString());
					}
					//errors.Add(new Tuple<string, object, object>(adr, fr[cell], sheet.Cells[adr].Value));
				}
			}
			catch (Exception e)
			{
				fileErr.WriteLine("Exception cell " + sheet.Name + "!" + adr + "\t" + fr[cell].ToString() + "\t" + sheet.Cells[adr].Value + "\t" + e.Message);
				fileErr.WriteLine("***************************");
				fileErr.WriteLine(e.ToString());
				nErrors++;
			}
		}

		fileErr.Close();
		return nErrors.ToString();
	}
}
