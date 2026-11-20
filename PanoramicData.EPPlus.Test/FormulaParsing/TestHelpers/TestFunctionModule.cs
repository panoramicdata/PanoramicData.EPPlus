using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

/// <summary>A functions module holding one custom function and its custom compiler.</summary>
public class TestFunctionModule : FunctionsModule
{
	public TestFunctionModule(ParsingContext context)
	{
		var myFunction = new MyFunction();
		var customCompiler = new MyFunctionCompiler(myFunction, context);
		Functions.Add(MyFunction.Name, myFunction);
		CustomCompilers.Add(typeof(MyFunction), customCompiler);
	}
}

public class MyFunction : ExcelFunction
{
	public static string Name => "MyFunction";
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context) => throw new NotImplementedException();
}

public class MyFunctionCompiler : FunctionCompiler
{
	public MyFunctionCompiler(MyFunction function, ParsingContext context) : base(function, context) { }
	public override CompileResult Compile(IEnumerable<Expression> children) => throw new NotImplementedException();
}