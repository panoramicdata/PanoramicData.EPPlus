using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Globalization;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	//[Ignore]
	//[TestMethod]
	public void ReadWorkSheet()
	{
		FileStream instream = new(_worksheetPath + @"Worksheet.xlsx", FileMode.Open, FileAccess.ReadWrite);
		using (ExcelPackage pck = new(instream))
		{
			var ws = pck.Workbook.Worksheets["Perf"];
			Assert.AreEqual("B5+B6", ws.Cells["H6"].Formula);

			ws = pck.Workbook.Worksheets["Comment"];
			var comment = ws.Cells["B2"].Comment;

			Assert.AreNotEqual(null, comment);
			Assert.AreEqual("Jan Källman", comment.Author);
			ws = pck.Workbook.Worksheets["Hidden"];
			Assert.AreEqual(eWorkSheetHidden.Hidden, ws.Hidden);

			ws = pck.Workbook.Worksheets["VeryHidden"];
			Assert.AreEqual(eWorkSheetHidden.VeryHidden, ws.Hidden);

			ws = pck.Workbook.Worksheets["RichText"];
			Assert.AreEqual("Room 02 & 03", ws.Cells["G1"].RichText.Text);

			_ = pck.Workbook.Worksheets["HeaderImage"];

			ws = pck.Workbook.Worksheets["newsheet"];
			Assert.AreEqual(true, ws.Cells["F2"].Style.Font.UnderLine);
			Assert.AreEqual(ExcelUnderLineType.Double, ws.Cells["F2"].Style.Font.UnderLineType);
			Assert.AreEqual(ExcelUnderLineType.SingleAccounting, ws.Cells["F3"].Style.Font.UnderLineType);
			Assert.AreEqual(ExcelUnderLineType.None, ws.Cells["F5"].Style.Font.UnderLineType);
			Assert.AreEqual(false, ws.Cells["F5"].Style.Font.UnderLine);

			Assert.AreEqual(ws.Cells["T20"].GetValue<string>(), 0.39618055555555554d.ToString(CultureInfo.CurrentCulture));
			Assert.AreEqual(0, ws.Cells["T20"].GetValue<int>());
			Assert.AreEqual(0, ws.Cells["T20"].GetValue<int?>());
			Assert.AreEqual(0.39618055555555554d, ws.Cells["T20"].GetValue<double>());
			Assert.AreEqual(0.39618055555555554d, ws.Cells["T20"].GetValue<double?>());
			Assert.AreEqual(0.396180555555556m, ws.Cells["T20"].GetValue<decimal>());
			Assert.AreEqual(0.396180555555556m, ws.Cells["T20"].GetValue<decimal?>());
			Assert.AreEqual(true, ws.Cells["T20"].GetValue<bool>());
			Assert.AreEqual(true, ws.Cells["T20"].GetValue<bool?>());
			Assert.AreEqual(ws.Cells["T20"].GetValue<DateTime>(), new DateTime(1899, 12, 30, 9, 30, 30));
			Assert.AreEqual(ws.Cells["T20"].GetValue<DateTime?>(), new DateTime(1899, 12, 30, 9, 30, 30));
			Assert.AreEqual(ws.Cells["T20"].GetValue<TimeSpan>(), new TimeSpan(693593, 9, 30, 30));
			Assert.AreEqual(ws.Cells["T20"].GetValue<TimeSpan?>(), new TimeSpan(693593, 9, 30, 30));
			Assert.AreEqual("09:30:30", ws.Cells["T20"].Text);

			Assert.AreEqual(ws.Cells["T24"].GetValue<string>(), 1.3961805555555555d.ToString(CultureInfo.CurrentCulture));
			Assert.AreEqual(1, ws.Cells["T24"].GetValue<int>());
			Assert.AreEqual(1, ws.Cells["T24"].GetValue<int?>());
			Assert.AreEqual(1.3961805555555555d, ws.Cells["T24"].GetValue<double>());
			Assert.AreEqual(1.3961805555555555d, ws.Cells["T24"].GetValue<double?>());
			Assert.AreEqual(1.39618055555556m, ws.Cells["T24"].GetValue<decimal>());
			Assert.AreEqual(1.39618055555556m, ws.Cells["T24"].GetValue<decimal?>());
			Assert.AreEqual(true, ws.Cells["T24"].GetValue<bool>());
			Assert.AreEqual(true, ws.Cells["T24"].GetValue<bool?>());
			Assert.AreEqual(ws.Cells["T24"].GetValue<DateTime>(), new DateTime(1899, 12, 31, 9, 30, 30));
			Assert.AreEqual(ws.Cells["T24"].GetValue<DateTime?>(), new DateTime(1899, 12, 31, 9, 30, 30));
			Assert.AreEqual(ws.Cells["T24"].GetValue<TimeSpan>(), new TimeSpan(693593, 33, 30, 30));
			Assert.AreEqual(ws.Cells["T24"].GetValue<TimeSpan?>(), new TimeSpan(693593, 33, 30, 30));
			Assert.AreEqual("09:30:30", ws.Cells["T24"].Text);

			Assert.AreEqual("40179", ws.Cells["U20"].GetValue<string>());
			Assert.AreEqual(40179, ws.Cells["U20"].GetValue<int>());
			Assert.AreEqual(40179, ws.Cells["U20"].GetValue<int?>());
			Assert.AreEqual(40179d, ws.Cells["U20"].GetValue<double>());
			Assert.AreEqual(40179d, ws.Cells["U20"].GetValue<double?>());
			Assert.AreEqual(40179m, ws.Cells["U20"].GetValue<decimal>());
			Assert.AreEqual(40179m, ws.Cells["U20"].GetValue<decimal?>());
			Assert.AreEqual(true, ws.Cells["U20"].GetValue<bool>());
			Assert.AreEqual(true, ws.Cells["U20"].GetValue<bool?>());
			Assert.AreEqual(ws.Cells["U20"].GetValue<DateTime>(), new DateTime(2010, 1, 1));
			Assert.AreEqual(ws.Cells["U20"].GetValue<DateTime?>(), new DateTime(2010, 1, 1));
			Assert.AreEqual("2010-01-01", ws.Cells["U20"].Text);

			Assert.AreEqual("102", ws.Cells["V20"].GetValue<string>());
			Assert.AreEqual(102, ws.Cells["V20"].GetValue<int>());
			Assert.AreEqual(102, ws.Cells["V20"].GetValue<int?>());
			Assert.AreEqual(102d, ws.Cells["V20"].GetValue<double>());
			Assert.AreEqual(102d, ws.Cells["V20"].GetValue<double?>());
			Assert.AreEqual(102m, ws.Cells["V20"].GetValue<decimal>());
			Assert.AreEqual(102m, ws.Cells["V20"].GetValue<decimal?>());
			Assert.AreEqual(true, ws.Cells["V20"].GetValue<bool>());
			Assert.AreEqual(true, ws.Cells["V20"].GetValue<bool?>());
			Assert.AreEqual(ws.Cells["V20"].GetValue<DateTime>(), new DateTime(1900, 4, 11));
			Assert.AreEqual(ws.Cells["V20"].GetValue<DateTime?>(), new DateTime(1900, 4, 11));
			Assert.AreEqual(ws.Cells["V20"].Text,
				$"$102{CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator}00");

			Assert.AreEqual(null, ws.Cells["W20"].GetValue<string>());
			Assert.AreEqual(0, ws.Cells["W20"].GetValue<int>());
			Assert.AreEqual(null, ws.Cells["W20"].GetValue<int?>());
			Assert.AreEqual(0d, ws.Cells["W20"].GetValue<double>());
			Assert.AreEqual(null, ws.Cells["W20"].GetValue<double?>());
			Assert.AreEqual(0m, ws.Cells["W20"].GetValue<decimal>());
			Assert.AreEqual(null, ws.Cells["W20"].GetValue<decimal?>());
			Assert.AreEqual(false, ws.Cells["W20"].GetValue<bool>());
			Assert.AreEqual(null, ws.Cells["W20"].GetValue<bool?>());
			Assert.AreEqual(ws.Cells["W20"].GetValue<DateTime>(), DateTime.MinValue);
			Assert.AreEqual(null, ws.Cells["W20"].GetValue<DateTime?>());
			Assert.AreEqual(ws.Cells["W20"].GetValue<TimeSpan>(), TimeSpan.Zero);
			Assert.AreEqual(null, ws.Cells["W20"].GetValue<TimeSpan?>());
			Assert.AreEqual(ws.Cells["W20"].Text, string.Empty);

			Assert.AreEqual("True", ws.Cells["Y20"].GetValue<string>());
			Assert.AreEqual(1, ws.Cells["Y20"].GetValue<int>());
			Assert.AreEqual(1, ws.Cells["Y20"].GetValue<int?>());
			Assert.AreEqual(1d, ws.Cells["Y20"].GetValue<double>());
			Assert.AreEqual(1d, ws.Cells["Y20"].GetValue<double?>());
			Assert.AreEqual(1m, ws.Cells["Y20"].GetValue<decimal>());
			Assert.AreEqual(1m, ws.Cells["Y20"].GetValue<decimal?>());
			Assert.AreEqual(true, ws.Cells["Y20"].GetValue<bool>());
			Assert.AreEqual(true, ws.Cells["Y20"].GetValue<bool?>());

			Assert.IsInstanceOfType<InvalidCastException>(GetValueException<DateTime>(ws.Cells["Y20"]));
			Assert.IsInstanceOfType<InvalidCastException>(GetValueException<DateTime?>(ws.Cells["Y20"]));
			Assert.IsInstanceOfType<InvalidCastException>(GetValueException<TimeSpan>(ws.Cells["Y20"]));
			Assert.IsInstanceOfType<InvalidCastException>(GetValueException<TimeSpan?>(ws.Cells["Y20"]));

			Assert.AreEqual("Text2", ws.Cells["Z20"].GetValue<string>());
			Assert.IsInstanceOfType<FormatException>(GetValueException<int>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<int?>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<double>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<double?>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<decimal>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<decimal?>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<bool>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<bool?>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<DateTime>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<DateTime?>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<TimeSpan>(ws.Cells["Z20"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<TimeSpan?>(ws.Cells["Z20"]));
			Assert.AreEqual("Text2", ws.Cells["Z20"].Text);

			AssertReads210AsNumber(ws.Cells["X19"]);
			Assert.AreEqual(true, ws.Cells["X19"].GetValue<bool>());
			Assert.AreEqual(true, ws.Cells["X19"].GetValue<bool?>());
			// double is converted to DateTime and TimeSpan, but value is meaningless
			Assert.IsNull(GetValueException<DateTime>(ws.Cells["X19"]));
			Assert.IsNull(GetValueException<DateTime?>(ws.Cells["X19"]));
			Assert.IsNull(GetValueException<TimeSpan>(ws.Cells["X19"]));
			Assert.IsNull(GetValueException<TimeSpan?>(ws.Cells["X19"]));
			Assert.AreEqual("210", ws.Cells["X19"].Text);

			Assert.IsInstanceOfType<string>(ws.Cells["AA19"].Value);
			AssertReads210AsNumber(ws.Cells["AA19"]);
			Assert.IsInstanceOfType<FormatException>(GetValueException<bool>(ws.Cells["AA19"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<bool?>(ws.Cells["AA19"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<DateTime>(ws.Cells["AA19"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<DateTime?>(ws.Cells["AA19"]));
			Assert.AreEqual(TimeSpan.FromDays(210), ws.Cells["AA19"].GetValue<TimeSpan>());
			Assert.AreEqual(TimeSpan.FromDays(210), ws.Cells["AA19"].GetValue<TimeSpan?>());
			Assert.AreEqual("210", ws.Cells["AA19"].Text);

			// blank string produces null for nullable primitive types and conversion to non-nullable types throws FormatException
			Assert.IsInstanceOfType<string>(ws.Cells["AA18"].Value);
			Assert.IsInstanceOfType<FormatException>(GetValueException<int>(ws.Cells["AA18"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<double>(ws.Cells["AA18"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<decimal>(ws.Cells["AA18"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<bool>(ws.Cells["AA18"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<DateTime>(ws.Cells["AA18"]));
			Assert.IsInstanceOfType<FormatException>(GetValueException<TimeSpan>(ws.Cells["AA18"]));

			Assert.IsNull(GetValueException<int?>(ws.Cells["AA18"]));
			Assert.IsNull(GetValueException<double?>(ws.Cells["AA18"]));
			Assert.IsNull(GetValueException<decimal?>(ws.Cells["AA18"]));
			Assert.IsNull(GetValueException<bool?>(ws.Cells["AA18"]));
			Assert.IsNull(GetValueException<DateTime?>(ws.Cells["AA18"]));
			Assert.IsNull(GetValueException<TimeSpan?>(ws.Cells["AA18"]));
			Assert.AreEqual(" ", ws.Cells["AA18"].Text);
		}

		instream.Close();
	}

	[Ignore]
	[TestMethod]
	public void ReadStreamWithTemplateWorkSheet()
	{
		FileStream instream = new(_worksheetPath + @"Worksheet.xlsx", FileMode.Open, FileAccess.Read);
		MemoryStream stream = new();
		using (ExcelPackage pck = new(stream, instream))
		{
			var ws = pck.Workbook.Worksheets["Perf"];
			Assert.AreEqual("B5+B6", ws.Cells["H6"].Formula);

			ws = pck.Workbook.Worksheets["newsheet"];
			Assert.AreEqual(ws.GetValue<DateTime>(20, 21), new DateTime(2010, 1, 1));

			ws = pck.Workbook.Worksheets["Loaded DataTable"];
			Assert.AreEqual("Row1", ws.GetValue<string>(2, 1));
			Assert.AreEqual(1, ws.GetValue<int>(2, 2));
			Assert.AreEqual(true, ws.GetValue<bool>(2, 3));
			Assert.AreEqual(1.5, ws.GetValue<double>(2, 4));

			ws = pck.Workbook.Worksheets["RichText"];

			var r1 = ws.Cells["A1"].RichText[0];
			Assert.AreEqual("Test", r1.Text);
			Assert.AreEqual(true, r1.Bold);

			ws = pck.Workbook.Worksheets["Pic URL"];
			Assert.AreEqual("http://epplus.codeplex.com", ((ExcelPicture)ws.Drawings["Pic URI"]).Hyperlink.ToString());

			Assert.AreEqual("\b\t", pck.Workbook.Worksheets["Address"].GetValue<string>(40, 1));

			pck.SaveAs(new FileInfo(Path.Combine("Test", "Worksheet2.xlsx")));
		}

		instream.Close();
	}
	//[Ignore]
	//[TestMethod]
	public void ReadStreamSaveAsStream()
	{
		if (!File.Exists(_worksheetPath + @"Worksheet.xlsx"))
		{
			Assert.Inconclusive("Worksheet.xlsx does not exists");
		}

		FileStream instream = new(_worksheetPath + @"Worksheet.xlsx", FileMode.Open, FileAccess.ReadWrite);
		MemoryStream stream = new();
		using (ExcelPackage pck = new(instream))
		{
			var ws = pck.Workbook.Worksheets["Names"];
			Assert.AreEqual(1, ws.Names["FullCol"].Start.Row);
			Assert.AreEqual(ExcelPackage.MaxRows, ws.Names["FullCol"].End.Row);
			pck.SaveAs(stream);
		}

		instream.Close();
	}

	/// <summary>
	///     Get exception from thrown from <see cref="ExcelRangeBase.GetValue{T}"/>
	/// </summary>
	/// <typeparam name="T">
	///     Type of value to attempt getting from <paramref name="range"/>
	/// </typeparam>
	/// <param name="range">
	///     Mandatory, cell range from which to attempt getting typed value
	/// </param>
	/// <returns>
	///     null if none is thrown
	/// </returns>
	private static Exception GetValueException<T>(ExcelRangeBase range) => GetException(() => range.GetValue<T>());

	/// <summary>
	///     Get exception thrown by <paramref name="action"/> when invoked.
	/// </summary>
	/// <param name="action">
	///     Mandatory
	/// </param>
	/// <returns>
	///     null if none is thrown
	/// </returns>
	private static Exception GetException(Action action)
	{
		try
		{
			action.Invoke();
			return null;
		}
		catch (Exception exception)
		{
			return exception;
		}
	}

	private static void AssertReads210AsNumber(ExcelRange cell)
	{
		Assert.AreEqual("210", cell.GetValue<string>());
		Assert.AreEqual(210, cell.GetValue<int>());
		Assert.AreEqual(210, cell.GetValue<int?>());
		Assert.AreEqual(210d, cell.GetValue<double>());
		Assert.AreEqual(210d, cell.GetValue<double?>());
		Assert.AreEqual(210m, cell.GetValue<decimal>());
		Assert.AreEqual(210m, cell.GetValue<decimal?>());
	}
}
