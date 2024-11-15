﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.ExcelDataProviderTests;

[TestClass]
public class ExcelDataProviderIntegrationTests
{
	private static ExcelCell CreateItem(object val, int row) => new(val, null, 0, row);



	//[TestMethod]
	//public void ShouldExecuteFormulaInRange()
	//{
	//    var expectedAddres = "A1:A2";
	//    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
	//    provider
	//        .Stub(x => x.GetRangeValues(expectedAddres))
	//        .Return(new object[] { 1, new ExcelCell(null, "SUM(1,2)", 0, 1) });
	//    var parser = new FormulaParser(provider);
	//    var result = parser.Parse(string.Format("sum({0})", expectedAddres));
	//    Assert.AreEqual(4d, result);
	//}

	//[TestMethod, ExpectedException(typeof(CircularReferenceException))]
	//public void ShouldHandleCircularReference2()
	//{
	//    var expectedAddres = "A1:A2";
	//    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
	//    provider
	//        .Stub(x => x.GetRangeValues(expectedAddres))
	//        .Return(new ExcelCell[] { CreateItem(1, 0), new ExcelCell(null, "SUM(A1:A2)",0, 1) });
	//    var parser = new FormulaParser(provider);
	//    var result = parser.Parse(string.Format("sum({0})", expectedAddres));
	//}
}
