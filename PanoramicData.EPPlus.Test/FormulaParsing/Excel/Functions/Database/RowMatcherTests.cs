using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using FakeItEasy;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Database;

[TestClass]
public class RowMatcherTests
{
	private static ExcelDatabaseCriteria GetCriteria(Dictionary<ExcelDatabaseCriteriaField, object> items)
	{
		var criteria = A.Fake<ExcelDatabaseCriteria>();

		A.CallTo(() => criteria.Items).Returns(items);
		return criteria;
	}

	private static bool IsMatch(object crit1, Dictionary<ExcelDatabaseCriteriaField, object> crit)
	{
		var data = new ExcelDatabaseRow();
		data["Crit1"] = crit1;
		data["Crit2"] = 2;
		data["Crit3"] = 3;

		var matcher = new RowMatcher();

		var criteria = GetCriteria(crit);

		return matcher.IsMatch(data, criteria);
	}
	[TestMethod]
	public void IsMatchShouldReturnTrueIfCriteriasMatch()
	{
		var crit = new Dictionary<ExcelDatabaseCriteriaField, object>
		{
			[new ExcelDatabaseCriteriaField("Crit1")] = 1,
			[new ExcelDatabaseCriteriaField("Crit3")] = 3
		};

		Assert.IsTrue(IsMatch(1, crit));
	}

	[TestMethod]
	public void IsMatchShouldReturnFalseIfCriteriasDoesNotMatch()
	{
		var crit = new Dictionary<ExcelDatabaseCriteriaField, object>
		{
			[new ExcelDatabaseCriteriaField("Crit1")] = 1,
			[new ExcelDatabaseCriteriaField("Crit3")] = 4
		};

		Assert.IsFalse(IsMatch(1, crit));
	}

	[TestMethod]
	public void IsMatchShouldMatchStrings1()
	{
		var crit = new Dictionary<ExcelDatabaseCriteriaField, object>
		{
			[new ExcelDatabaseCriteriaField("Crit1")] = "1",
			[new ExcelDatabaseCriteriaField("Crit3")] = 3
		};

		Assert.IsTrue(IsMatch("1", crit));
	}

	[TestMethod]
	public void IsMatchShouldMatchStrings2()
	{
		var crit = new Dictionary<ExcelDatabaseCriteriaField, object>
		{
			[new ExcelDatabaseCriteriaField("Crit1")] = "1",
			[new ExcelDatabaseCriteriaField("Crit3")] = 3
		};

		Assert.IsFalse(IsMatch("2", crit));
	}

	[TestMethod]
	public void IsMatchShouldMatchWildcardStrings()
	{
		var crit = new Dictionary<ExcelDatabaseCriteriaField, object>
		{
			[new ExcelDatabaseCriteriaField("Crit1")] = "t*t",
			[new ExcelDatabaseCriteriaField("Crit3")] = 3
		};

		Assert.IsTrue(IsMatch("test", crit));
	}

	[TestMethod]
	public void IsMatchShouldMatchNumericExpression()
	{
		var crit = new Dictionary<ExcelDatabaseCriteriaField, object>
		{
			[new ExcelDatabaseCriteriaField("Crit2")] = "<3",
			[new ExcelDatabaseCriteriaField("Crit3")] = 3
		};

		Assert.IsTrue(IsMatch("test", crit));
	}

	[TestMethod]
	public void IsMatchShouldHandleFieldIndex()
	{
		var crit = new Dictionary<ExcelDatabaseCriteriaField, object>
		{
			[new ExcelDatabaseCriteriaField(2)] = "<3",
			[new ExcelDatabaseCriteriaField("Crit3")] = 3
		};

		Assert.IsTrue(IsMatch("test", crit));
	}
}
