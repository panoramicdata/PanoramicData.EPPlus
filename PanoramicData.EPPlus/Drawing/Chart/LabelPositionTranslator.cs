namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Translates label positions to and from their chart XML values.
/// </summary>
internal static class LabelPositionTranslator
{
	internal static string ToText(eLabelPosition pos) => pos switch
	{
		eLabelPosition.Bottom => "b",
		eLabelPosition.Center => "ctr",
		eLabelPosition.InBase => "inBase",
		eLabelPosition.InEnd => "inEnd",
		eLabelPosition.Left => "l",
		eLabelPosition.Right => "r",
		eLabelPosition.Top => "t",
		eLabelPosition.OutEnd => "outEnd",
		_ => "bestFit",
	};

	internal static eLabelPosition ToEnum(string pos) => pos switch
	{
		"b" => eLabelPosition.Bottom,
		"ctr" => eLabelPosition.Center,
		"inBase" => eLabelPosition.InBase,
		"inEnd" => eLabelPosition.InEnd,
		"l" => eLabelPosition.Left,
		"r" => eLabelPosition.Right,
		"t" => eLabelPosition.Top,
		"outEnd" => eLabelPosition.OutEnd,
		_ => eLabelPosition.BestFit,
	};
}
