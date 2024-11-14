using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.ErrorHandling;

[TestClass]
public class FormulaErrorHandlingTestBase
{
	protected ExcelPackage Package;
	protected ExcelWorksheet Worksheet;

	public void BaseInitialize()
	{
#if !Core
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#else
		var dir = AppContext.BaseDirectory;
#endif
		var Package = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "FormulaTest.xlsx")));
		Worksheet = Package.Workbook.Worksheets["ValidateFormulas"];
		Package.Workbook.Calculate();
	}

	public void BaseCleanup() => Package.Dispose();
}
