﻿using System;
using System.Globalization;
using System.Xml;
using System.Drawing;

namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfStyleConditionalFormatting : DxfStyleBase<ExcelDxfStyleConditionalFormatting>
{
	readonly XmlHelperInstance _helper;
	internal ExcelDxfStyleConditionalFormatting(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) : base(styles)
	{
		NumberFormat = new ExcelDxfNumberFormat(_styles);
		Font = new ExcelDxfFontBase(_styles);
		Border = new ExcelDxfBorderBase(_styles);
		Fill = new ExcelDxfFill(_styles);
		if (topNode != null)
		{
			_helper = new XmlHelperInstance(nameSpaceManager, topNode);
			NumberFormat.NumFmtID = _helper.GetXmlNodeInt("d:numFmt/@numFmtId");
			NumberFormat.Format = _helper.GetXmlNodeString("d:numFmt/@formatCode");
			if (NumberFormat.NumFmtID < 164 && string.IsNullOrEmpty(NumberFormat.Format))
			{
				NumberFormat.Format = ExcelNumberFormat.GetFromBuildInFromID(NumberFormat.NumFmtID);
			}

			Font.Bold = _helper.GetXmlNodeBoolNullable("d:font/d:b/@val");
			Font.Italic = _helper.GetXmlNodeBoolNullable("d:font/d:i/@val");
			Font.Strike = _helper.GetXmlNodeBoolNullable("d:font/d:strike");
			Font.Underline = GetUnderLineEnum(_helper.GetXmlNodeString("d:font/d:u/@val"));
			Font.Color = GetColor(_helper, "d:font/d:color");

			Border.Left = GetBorderItem(_helper, "d:border/d:left");
			Border.Right = GetBorderItem(_helper, "d:border/d:right");
			Border.Bottom = GetBorderItem(_helper, "d:border/d:bottom");
			Border.Top = GetBorderItem(_helper, "d:border/d:top");

			Fill.PatternType = GetPatternTypeEnum(_helper.GetXmlNodeString("d:fill/d:patternFill/@patternType"));
			Fill.BackgroundColor = GetColor(_helper, "d:fill/d:patternFill/d:bgColor/");
			Fill.PatternColor = GetColor(_helper, "d:fill/d:patternFill/d:fgColor/");
		}
		else
		{
			_helper = new XmlHelperInstance(nameSpaceManager);
		}

		_helper.SchemaNodeOrder = ["font", "numFmt", "fill", "border"];
	}
	private ExcelDxfBorderItem GetBorderItem(XmlHelperInstance helper, string path)
	{
		ExcelDxfBorderItem bi = new(_styles)
		{
			Style = GetBorderStyleEnum(helper.GetXmlNodeString(path + "/@style")),
			Color = GetColor(helper, path + "/d:color")
		};
		return bi;
	}
	private static ExcelBorderStyle GetBorderStyleEnum(string style)
	{
		if (style == "") return ExcelBorderStyle.None;
		var sInStyle = style[..1].ToUpper(CultureInfo.InvariantCulture) + style[1..];
		try
		{
			return Enum.Parse<ExcelBorderStyle>(sInStyle);
		}
		catch
		{
			return ExcelBorderStyle.None;
		}

	}
	private static ExcelFillStyle GetPatternTypeEnum(string patternType)
	{
		if (patternType == "") return ExcelFillStyle.None;
		patternType = patternType[..1].ToUpper(CultureInfo.InvariantCulture) + patternType[1..];
		try
		{
			return Enum.Parse<ExcelFillStyle>(patternType);
		}
		catch
		{
			return ExcelFillStyle.None;
		}
	}
	private ExcelDxfColor GetColor(XmlHelperInstance helper, string path)
	{
		ExcelDxfColor ret = new(_styles)
		{
			Theme = helper.GetXmlNodeIntNull(path + "/@theme"),
			Index = helper.GetXmlNodeIntNull(path + "/@indexed")
		};
		var rgb = helper.GetXmlNodeString(path + "/@rgb");
		if (rgb != "")
		{
			ret.Color = Color.FromArgb(int.Parse(rgb[..2], NumberStyles.AllowHexSpecifier),
										int.Parse(rgb.Substring(2, 2), NumberStyles.AllowHexSpecifier),
										int.Parse(rgb.Substring(4, 2), NumberStyles.AllowHexSpecifier),
										int.Parse(rgb.Substring(6, 2), NumberStyles.AllowHexSpecifier));
		}

		ret.Auto = helper.GetXmlNodeBoolNullable(path + "/@auto");
		ret.Tint = helper.GetXmlNodeDoubleNull(path + "/@tint");
		return ret;
	}
	private static ExcelUnderLineType? GetUnderLineEnum(string value) => value.ToLower(CultureInfo.InvariantCulture) switch
	{
		"single" => (ExcelUnderLineType?)ExcelUnderLineType.Single,
		"double" => (ExcelUnderLineType?)ExcelUnderLineType.Double,
		"singleaccounting" => (ExcelUnderLineType?)ExcelUnderLineType.SingleAccounting,
		"doubleaccounting" => (ExcelUnderLineType?)ExcelUnderLineType.DoubleAccounting,
		_ => null,
	};
	internal int DxfId { get; set; }
	public ExcelDxfFontBase Font { get; set; }
	public ExcelDxfNumberFormat NumberFormat { get; set; }
	public ExcelDxfFill Fill { get; set; }
	public ExcelDxfBorderBase Border { get; set; }
	protected internal override string Id => NumberFormat.Id + Font.Id + Border.Id + Fill.Id +
				(AllowChange ? "" : DxfId.ToString());//If allowchange is false we add the dxfID to ensure it's not used when conditional formatting is updated);
	protected internal override ExcelDxfStyleConditionalFormatting Clone()
	{
		var s = new ExcelDxfStyleConditionalFormatting(_helper.NameSpaceManager, null, _styles)
		{
			Font = Font.Clone(),
			NumberFormat = NumberFormat.Clone(),
			Fill = Fill.Clone(),
			Border = Border.Clone()
		};
		return s;
	}

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		if (Font.HasValue) Font.CreateNodes(helper, "d:font");
		if (NumberFormat.HasValue) NumberFormat.CreateNodes(helper, "d:numFmt");
		if (Fill.HasValue) Fill.CreateNodes(helper, "d:fill");
		if (Border.HasValue) Border.CreateNodes(helper, "d:border");
	}
	protected internal override bool HasValue => Font.HasValue || NumberFormat.HasValue || Fill.HasValue || Border.HasValue;
}
