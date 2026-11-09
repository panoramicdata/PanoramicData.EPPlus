
namespace PanoramicData.EPPlus.Test.FormulaParsing;

public partial class FormulaR1C1Tests
{
	[TestMethod]
	public void RCRelativeToAB()
	{
		var fR1C1 = "SUMIFS(C,C2,RC1)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		Assert.AreEqual("SUMIFS(C:C,$B:$B,$A5)", f);
	}
	[TestMethod]
	public void RRelativeToAB()
	{
		var fR1C1 = "SUMIFS(R,C2,RC1)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		Assert.AreEqual("SUMIFS(5:5,$B:$B,$A5)", f);
	}
	[TestMethod]
	public void RCRelativeToABToR1C1()
	{
		var fR1C1 = "SUMIFS(C,C2,RC1)";
		_sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 3].Formula;
		_sheet.Cells[5, 3].Formula = f;
		Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
	}
	[TestMethod]
	public void RCRelativeToABToR1C1_2()
	{
		var fR1C1 = "SUM(RC9:RC[-1])";
		_sheet.Cells[5, 13].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[5, 13].Formula;
		Assert.AreEqual("SUM($I5:L5)", f);
		_sheet.Cells[5, 13].Formula = f;
		Assert.AreEqual(fR1C1, _sheet.Cells[5, 13].FormulaR1C1);
	}
	[TestMethod]
	public void RCFixToABToR1C1_2()
	{
		var fR1C1 = "RC28-SUM(RC12:RC21)";
		_sheet.Cells[6, 13].FormulaR1C1 = fR1C1;
		var f = _sheet.Cells[6, 13].Formula;
		Assert.AreEqual("$AB6-SUM($L6:$U6)", f);
		_sheet.Cells[6, 13].Formula = f;
		Assert.AreEqual(fR1C1, _sheet.Cells[6, 13].FormulaR1C1);
	}
	[TestMethod]
	public void SimpleRelativeR1C1()
	{
		var fR1C1 = "R[-1]C[-5]";
		var c = _sheet.Cells[7, 7];
		c.FormulaR1C1 = fR1C1;
		var f = c.Formula;
		Assert.AreEqual("B6", f);
		c.Formula = f;
		Assert.AreEqual(fR1C1, c.FormulaR1C1);
	}
	[TestMethod]
	public void SimpleAbsR1C1()
	{
		var fR1C1 = "R1C5";
		var c = _sheet.Cells[8, 8];
		c.FormulaR1C1 = fR1C1;
		var f = c.Formula;
		Assert.AreEqual("$E$1", f);
		c.Formula = f;
		Assert.AreEqual(fR1C1, c.FormulaR1C1);
	}
	[TestMethod]
	public void FullTwoColumn()
	{
		var formula = "VLOOKUP(C2,A:B,1,0)";
		var c = _sheet.Cells["D2"];
		c.Formula = formula;
		Assert.AreEqual("VLOOKUP(RC[-1],C[-3]:C[-2],1,0)", c.FormulaR1C1);
		var r1c1 = c.FormulaR1C1;
		c.FormulaR1C1 = r1c1;
		Assert.AreEqual(c.Formula, formula);
	}
	[TestMethod]
	public void FullColumn()
	{
		var formula = "VLOOKUP(C2,A:A,1,0)";
		var c = _sheet.Cells["D2"];
		c.Formula = formula;
		Assert.AreEqual("VLOOKUP(RC[-1],C[-3],1,0)", c.FormulaR1C1);
		var r1c1 = c.FormulaR1C1;
		c.FormulaR1C1 = r1c1;
		Assert.AreEqual(c.Formula, formula);
	}
	[TestMethod]
	public void FullTwoRow()
	{
		var formula = "VLOOKUP(C3,1:2,1,0)";
		var c = _sheet.Cells["D3"];
		c.Formula = formula;
		Assert.AreEqual("VLOOKUP(RC[-1],R[-2]:R[-1],1,0)", c.FormulaR1C1);
		var r1c1 = c.FormulaR1C1;
		c.FormulaR1C1 = r1c1;
		Assert.AreEqual(c.Formula, formula);
	}
	[TestMethod]
	public void FullRow()
	{
		var formula = "VLOOKUP(C3,1:1,1,0)";
		var c = _sheet.Cells["D3"];
		c.Formula = formula;
		Assert.AreEqual("VLOOKUP(RC[-1],R[-2],1,0)", c.FormulaR1C1);
		var r1c1 = c.FormulaR1C1;
		c.FormulaR1C1 = r1c1;
		Assert.AreEqual(c.Formula, formula);
	}

	[TestMethod]
	public void OutOfRangeCol()
	{
		_sheet.Cells["a3"].FormulaR1C1 = "R[-3]C";
		Assert.AreEqual("#REF!", _sheet.Cells["a3"].Formula);
		_sheet.Cells["a3"].FormulaR1C1 = "R[-2]C";
		Assert.AreEqual("A1", _sheet.Cells["a3"].Formula);

		_sheet.Cells["B3"].FormulaR1C1 = "RC[-2]";
		Assert.AreEqual("#REF!", _sheet.Cells["B3"].Formula);
		_sheet.Cells["B3"].FormulaR1C1 = "RC[-1]";
		Assert.AreEqual("A3", _sheet.Cells["B3"].Formula);

	}
}
