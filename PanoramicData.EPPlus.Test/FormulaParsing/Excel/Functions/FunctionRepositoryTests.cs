using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public class FunctionRepositoryTests
{
	#region LoadModule Tests
	[TestMethod]
	public void LoadModulePopulatesFunctionsAndCustomCompilers()
	{
		var functionRepository = FunctionRepository.Create();
		Assert.IsFalse(functionRepository.IsFunctionName(MyFunction.Name));
		Assert.IsFalse(functionRepository.CustomCompilers.ContainsKey(typeof(MyFunction)));
		functionRepository.LoadModule(new TestFunctionModule(ParsingContext.Create()));
		Assert.IsTrue(functionRepository.IsFunctionName(MyFunction.Name));
		Assert.IsTrue(functionRepository.CustomCompilers.ContainsKey(typeof(MyFunction)));
		// Make sure reloading the module overwrites previous functions and compilers
		functionRepository.LoadModule(new TestFunctionModule(ParsingContext.Create()));
	}
	#endregion
}
