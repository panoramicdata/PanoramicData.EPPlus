using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.ExpressionGraph.FunctionCompilers;

[TestClass]
public class FunctionCompilerFactoryTests
{
	private ParsingContext _context;

	[TestInitialize]
	public void Initialize() => _context = ParsingContext.Create();
	#region Create Tests
	[TestMethod]
	public void CreateHandlesStandardFunctionCompiler()
	{
		var functionRepository = FunctionRepository.Create();
		var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
		var function = new Sum();
		var functionCompiler = functionCompilerFactory.Create(function);
		Assert.IsInstanceOfType<DefaultCompiler>(functionCompiler);
	}

	[TestMethod]
	public void CreateHandlesSpecialIfCompiler()
	{
		var functionRepository = FunctionRepository.Create();
		var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
		var function = new If();
		var functionCompiler = functionCompilerFactory.Create(function);
		Assert.IsInstanceOfType<IfFunctionCompiler>(functionCompiler);
	}

	[TestMethod]
	public void CreateHandlesSpecialIfErrorCompiler()
	{
		var functionRepository = FunctionRepository.Create();
		var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
		var function = new IfError();
		var functionCompiler = functionCompilerFactory.Create(function);
		Assert.IsInstanceOfType<IfErrorFunctionCompiler>(functionCompiler);
	}

	[TestMethod]
	public void CreateHandlesSpecialIfNaCompiler()
	{
		var functionRepository = FunctionRepository.Create();
		var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
		var function = new IfNa();
		var functionCompiler = functionCompilerFactory.Create(function);
		Assert.IsInstanceOfType<IfNaFunctionCompiler>(functionCompiler);
	}

	[TestMethod]
	public void CreateHandlesLookupFunctionCompiler()
	{
		var functionRepository = FunctionRepository.Create();
		var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
		var function = new Column();
		var functionCompiler = functionCompilerFactory.Create(function);
		Assert.IsInstanceOfType<LookupFunctionCompiler>(functionCompiler);
	}

	[TestMethod]
	public void CreateHandlesErrorFunctionCompiler()
	{
		var functionRepository = FunctionRepository.Create();
		var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
		var function = new IsError();
		var functionCompiler = functionCompilerFactory.Create(function);
		Assert.IsInstanceOfType<ErrorHandlingFunctionCompiler>(functionCompiler);
	}

	[TestMethod]
	public void CreateHandlesCustomFunctionCompiler()
	{
		var functionRepository = FunctionRepository.Create();
		functionRepository.LoadModule(new TestFunctionModule(_context));
		var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
		var function = new MyFunction();
		var functionCompiler = functionCompilerFactory.Create(function);
		Assert.IsInstanceOfType<MyFunctionCompiler>(functionCompiler);
	}
	#endregion

	#region Nested Classes
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
		public const string Name = "MyFunction";
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context) => throw new NotImplementedException();
	}

	public class MyFunctionCompiler : FunctionCompiler
	{
		public MyFunctionCompiler(MyFunction function, ParsingContext context) : base(function, context) { }
		public override CompileResult Compile(IEnumerable<Expression> children) => throw new NotImplementedException();
	}
	#endregion
}
