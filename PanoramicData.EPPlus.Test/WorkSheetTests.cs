
namespace PanoramicData.EPPlus.Test;

internal class BaseItem
{
	public int Ref { get; set; }
	public string BaseProp1 { get; set; }
}

internal class InhItem : BaseItem
{
	public int TophRef { get; set; }
	public string TopProp1 { get; set; }
}

[TestClass]
public partial class WorkSheetTest : TestBase
{
	[TestMethod]
	public void RunWorksheetTests()
	{
		InsertDeleteTestRows();
		InsertDeleteTestColumns();
		LoadData();
		AutoFilter();
		StyleFill();
		Performance();
		RichTextCells();
		TestComments();
		Hyperlink();
		PictureURL();
		CopyOverwrite();
		HideTest();
		VeryHideTest();
		PrinterSettings();
		Address();
		Merge();
		Encoding();
		LoadText();
		LoadDataReader();
		LoadDataTable();
		LoadFromCollectionTest();
		LoadFromEmptyCollectionTest();
		LoadArray();
		WorksheetCopy();
		DefaultColWidth();
		CopyTable();
		AutoFitColumns();
		CopyRange();
		CopyMergedRange();
		ValueError();
		FormulaOverwrite();
		FormulaError();
		StyleNameTest();
		NamedStyles();
		TableTest();
		DefinedName();
		CreatePivotTable();
		AddChartSheet();
		SetHeaderFooterImage();

		SaveWorksheet("Worksheet.xlsx");

		ReadWorkSheet();
		ReadStreamSaveAsStream();
	}
}
