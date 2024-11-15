﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public class ArgumentParsersTests
{
	[TestMethod]
	public void ShouldReturnSameInstanceOfIntParserWhenCalledTwice()
	{
		var parsers = new ArgumentParsers();
		var parser1 = parsers.GetParser(DataType.Integer);
		var parser2 = parsers.GetParser(DataType.Integer);
		Assert.AreEqual(parser1, parser2);
	}
}
