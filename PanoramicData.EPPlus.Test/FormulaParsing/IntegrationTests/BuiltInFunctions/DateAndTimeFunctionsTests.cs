using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public partial class DateAndTimeFunctionsTests : FormulaParserTestBase
{
	[TestInitialize]
	public void Setup()
	{
		var excelDataProvider = A.Fake<ExcelDataProvider>();
		_parser = new FormulaParser(excelDataProvider);
	}
#if (!Core)
        [TestMethod]
        public void DateValueTest1()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "21 JAN 2015";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(2015, 1, 21).ToOADate(), ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void DateValueTestWithoutYear()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "21 JAN";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(currentYear, 1, 21).ToOADate(), ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void DateValueTestWithTwoDigitYear()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var expectedYear = 1930;
            ws.Cells["A1"].Value = "01/01/30";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(expectedYear, 1, 1).ToOADate(), ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void DateValueTestWithTwoDigitYear2()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var expectedYear = 2029;
            ws.Cells["A1"].Value = "01/01/29";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(expectedYear, 1, 1).ToOADate(), ws.Cells["B1"].Value);
        }


        [TestMethod]
        public void TimeValueTestPm()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "2:23 pm";
            ws.Cells["B1"].Formula = "TimeValue(A1)";
            ws.Calculate();
            var result = (double) ws.Cells["B1"].Value;
            Assert.AreEqual(0.599, Math.Round(result, 3));
        }


        [TestMethod]
        public void TimeValueTestFullDate()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "01/01/2011 02:23";
            ws.Cells["B1"].Formula = "TimeValue(A1)";
            ws.Calculate();
            var result = (double)ws.Cells["B1"].Value;
            Assert.AreEqual(0.099, Math.Round(result, 3));
        }
#endif
}
