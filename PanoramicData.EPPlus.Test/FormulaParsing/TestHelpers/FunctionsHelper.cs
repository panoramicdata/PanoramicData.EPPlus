using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System.Collections.Generic;

namespace PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

public static class FunctionsHelper
{
	public static IEnumerable<FunctionArgument> CreateArgs(params object?[] args)
	{
		var list = new List<FunctionArgument>();
		foreach (var arg in args)
		{
			list.Add(new FunctionArgument(arg));
		}

		return list;
	}

	public static IEnumerable<FunctionArgument> Empty() => new List<FunctionArgument>() { new(null) };
}
