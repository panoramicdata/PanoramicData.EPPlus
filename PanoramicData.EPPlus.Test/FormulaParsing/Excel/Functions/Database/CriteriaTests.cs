using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Database;

[TestClass]
public class CriteriaTests
{
	private static ExcelDatabaseCriteria ReadCriteria(params (string Address, object Value)[] cells)
	{
		using var package = new ExcelPackage();
		var sheet = package.Workbook.Worksheets.Add("test");
		foreach (var (address, value) in cells)
		{
			sheet.Cells[address].Value = value;
		}

		var provider = new EpplusExcelDataProvider(package);
		return new ExcelDatabaseCriteria(provider, "A1:B2");
	}

	[TestMethod]
	public void CriteriaShouldReadFieldsAndValues()
	{
		var criteria = ReadCriteria(("A1", "Crit1"), ("B1", "Crit2"), ("A2", 1), ("B2", 2));

		Assert.AreEqual(2, criteria.Items.Count);
		Assert.AreEqual("crit1", criteria.Items.Keys.First().ToString());
		Assert.AreEqual("crit2", criteria.Items.Keys.Last().ToString());
		Assert.AreEqual(1, criteria.Items.Values.First());
		Assert.AreEqual(2, criteria.Items.Values.Last());
	}

	[TestMethod]
	public void CriteriaShouldIgnoreEmptyFields1()
	{
		var criteria = ReadCriteria(("A1", "Crit1"), ("B1", "Crit2"), ("A2", 1));

		Assert.AreEqual(1, criteria.Items.Count);
		Assert.AreEqual("crit1", criteria.Items.Keys.First().ToString());
		Assert.AreEqual(1, criteria.Items.Values.Last());
	}

	[TestMethod]
	public void CriteriaShouldIgnoreEmptyFields2()
	{
		var criteria = ReadCriteria(("A1", "Crit1"), ("A2", 1));

		Assert.AreEqual(1, criteria.Items.Count);
		Assert.AreEqual("crit1", criteria.Items.Keys.First().ToString());
		Assert.AreEqual(1, criteria.Items.Values.Last());
	}
}