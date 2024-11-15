﻿using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public class ExcelFunctionTests
{
	private class ExcelFunctionTester : ExcelFunction
	{
		public IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerableImpl(IEnumerable<FunctionArgument> args) => ArgsToDoubleEnumerable(args, ParsingContext.Create());
		#region Other members
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context) => throw new NotImplementedException();
		#endregion
	}

	[TestMethod]
	public void ArgsToDoubleEnumerableShouldHandleInnerEnumerables()
	{
		var args = FunctionsHelper.CreateArgs(1, 2, FunctionsHelper.CreateArgs(3, 4));
		var tester = new ExcelFunctionTester();
		var result = tester.ArgsToDoubleEnumerableImpl(args);
		Assert.AreEqual(4, result.Count());
	}
}
