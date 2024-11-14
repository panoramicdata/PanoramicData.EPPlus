namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfBorderItem : DxfStyleBase<ExcelDxfBorderItem>
{
	internal ExcelDxfBorderItem(ExcelStyles styles) :
		base(styles)
	{
		Color = new ExcelDxfColor(styles);
	}
	public ExcelBorderStyle? Style { get; set; }
	public ExcelDxfColor Color { get; internal set; }

	protected internal override string Id => GetAsString(Style) + "|" + (Color == null ? "" : Color.Id);

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		SetValueEnum(helper, path + "/@style", Style);
		SetValueColor(helper, path + "/d:color", Color);
	}
	protected internal override bool HasValue => Style != null || Color.HasValue;
	protected internal override ExcelDxfBorderItem Clone() => new(_styles) { Style = Style, Color = Color };
}
