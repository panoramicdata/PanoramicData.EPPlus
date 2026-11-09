using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

[TestClass]
public class LookupNavigatorTests
{
	const string WorksheetName = "";
	private static LookupArguments GetArgs(params object[] args)
	{
		var lArgs = FunctionsHelper.CreateArgs(args);
		return new LookupArguments(lArgs, ParsingContext.Create());
	}

	private static ParsingContext GetContext(ExcelDataProvider provider)
	{
		var ctx = ParsingContext.Create();
		ctx.Scopes.NewScope(new RangeAddress { Worksheet = WorksheetName, FromCol = 1, FromRow = 1 });
		ctx.ExcelDataProvider = provider;
		return ctx;
	}

	private static LookupNavigator CreateVerticalNavigator(LookupArguments args, ExcelCellAddress? dimensionEnd, params (int Row, int Col, object Value)[] cells)
	{
		var provider = A.Fake<ExcelDataProvider>();
		if (dimensionEnd != null)
		{
			A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(dimensionEnd);
		}

		foreach (var (row, col, value) in cells)
		{
			A.CallTo(() => provider.GetCellValue(WorksheetName, row, col)).Returns(value);
		}

		return LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
	}

	[TestMethod]
	public void CurrentValueShouldBeFirstCell()
	{
		var navigator = CreateVerticalNavigator(GetArgs(3, "A1:B2", 1), null, (1, 1, 3), (2, 1, 4));
		Assert.AreEqual(3, navigator.CurrentValue);
	}

	[TestMethod]
	public void MoveNextShouldReturnFalseIfLastCell()
	{
		var navigator = CreateVerticalNavigator(GetArgs(3, "A1:B1", 1), null, (1, 1, 3), (2, 1, 4));
		Assert.IsFalse(navigator.MoveNext());
	}

	[TestMethod]
	public void HasNextShouldBeTrueIfNotLastCell()
	{
		var navigator = CreateVerticalNavigator(GetArgs(3, "A1:B2", 1), new ExcelCellAddress(5, 5), (1, 1, 3), (2, 1, 4));
		Assert.IsTrue(navigator.MoveNext());
	}

	[TestMethod]
	public void MoveNextShouldNavigateVertically()
	{
		var navigator = CreateVerticalNavigator(GetArgs(6, "A1:B2", 1), new ExcelCellAddress(100, 10), (1, 1, 3), (2, 1, 4));
		navigator.MoveNext();
		Assert.AreEqual(4, navigator.CurrentValue);
	}

	[TestMethod]
	public void MoveNextShouldIncreaseIndex()
	{
		var navigator = CreateVerticalNavigator(GetArgs(6, "A1:B2", 1), new ExcelCellAddress(5, 5), (1, 1, 3), (1, 2, 4));
		Assert.AreEqual(0, navigator.Index);
		navigator.MoveNext();
		Assert.AreEqual(1, navigator.Index);
	}

	[TestMethod]
	public void GetLookupValueShouldReturnCorrespondingValue()
	{
		var navigator = CreateVerticalNavigator(GetArgs(6, "A1:B2", 2), new ExcelCellAddress(5, 5), (1, 1, 3), (1, 2, 4));
		Assert.AreEqual(4, navigator.GetLookupValue());
	}

	[TestMethod]
	public void GetLookupValueShouldReturnCorrespondingValueWithOffset()
	{
		var args = new LookupArguments(3, "A1:A4", 3, 2, false, null);
		var navigator = CreateVerticalNavigator(args, new ExcelCellAddress(5, 5), (1, 1, 3), (3, 3, 4));
		Assert.AreEqual(4, navigator.GetLookupValue());
	}
}