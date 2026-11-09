/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System.Diagnostics.CodeAnalysis;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Utilities;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public abstract class FunctionCompiler
{
	protected ExcelFunction Function
	{
		get;
		private set;
	}

	protected ParsingContext Context
	{
		get;
		private set;
	}

	[SuppressMessage("Major Code Smell", "S3442:\"abstract\" classes should not have \"public\" constructors", Justification = "Published API: narrowing this constructor to protected is a breaking change for reflection and API compatibility.")]
	public FunctionCompiler(ExcelFunction function, ParsingContext context)
	{
		Require.That(function).Named("function").IsNotNull();
		Require.That(context).Named("context").IsNotNull();
		Function = function;
		Context = context;
	}

	protected static void BuildFunctionArguments(CompileResult compileResult, DataType dataType, List<FunctionArgument> args)
	{
		if (compileResult.Result is IEnumerable<object> and not ExcelDataProvider.IRangeInfo)
		{
			var compileResultFactory = new CompileResultFactory();
			var argList = new List<FunctionArgument>();
			var objects = compileResult.Result as IEnumerable<object>;
			foreach (var arg in objects)
			{
				var cr = compileResultFactory.Create(arg);
				BuildFunctionArguments(cr, dataType, argList);
			}

			args.Add(new FunctionArgument(argList));
		}
		else
		{
			var funcArg = new FunctionArgument(compileResult.Result, dataType)
			{
				ExcelAddressReferenceId = compileResult.ExcelAddressReferenceId
			};
			args.Add(funcArg);
		}
	}

	protected static void BuildFunctionArguments(CompileResult result, List<FunctionArgument> args) => BuildFunctionArguments(result, result.DataType, args);

	private protected CompileResult CompileWithFallback(IEnumerable<Expression> children, System.Func<CompileResult, bool> useFallback)
	{
		var args = new List<FunctionArgument>();
		Function.BeforeInvoke(Context);
		var firstChild = children.First();
		var lastChild = children.ElementAt(1);
		try
		{
			var result = firstChild.Compile();
			args.Add(new FunctionArgument(useFallback(result) ? lastChild.Compile().Result : result.Result));
		}
		catch (Exceptions.ExcelErrorValueException)
		{
			args.Add(new FunctionArgument(lastChild.Compile().Result));
		}

		return Function.Execute(args, Context);
	}

	public abstract CompileResult Compile(IEnumerable<Expression> children);
}
