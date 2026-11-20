using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public sealed partial class ExcelStyles
{
	public ExcelNamedStyleXml CreateNamedStyle(string name) => CreateNamedStyle(name, null);
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	public ExcelNamedStyleXml CreateNamedStyle(string name, ExcelStyle Template)
	{
		if (_wb.Styles.NamedStyles.ExistsKey(name))
		{
			throw new Exception(string.Format("Key {0} already exists in collection", name));
		}

		ExcelNamedStyleXml style;
		style = new ExcelNamedStyleXml(NameSpaceManager, this);
		int xfIdCopy, positionID;
		ExcelStyles styles;
		if (Template == null)
		{
			xfIdCopy = 0;
			positionID = -1;
			styles = this;
		}
		else
		{
			if (Template.PositionID < 0 && Template.Styles == this)
			{
				xfIdCopy = Template.Index;

				positionID = Template.PositionID;
				styles = this;
			}
			else
			{
				xfIdCopy = Template.XfId;
				positionID = -1;
				styles = Template.Styles;
			}
		}
		//Clone namedstyle
		var styleXfId = CloneStyle(styles, xfIdCopy, true);
		//Close cells style
		CellStyleXfs[styleXfId].XfId = CellStyleXfs.Count - 1;
		var xfid = CloneStyle(styles, xfIdCopy, true, true); //Always add a new style (We create a new named style here)
		CellXfs[xfid].XfId = styleXfId;
		style.Style = new ExcelStyle(this, NamedStylePropertyChange, positionID, name, styleXfId);
		style.StyleXfId = styleXfId;

		style.Name = name;
		var ix = _wb.Styles.NamedStyles.Add(style.Name, style);
		style.Style.SetIndex(ix);
		return style;
	}
	internal int GetStyleIdFromName(string Name)
	{
		var i = NamedStyles.FindIndexByID(Name);
		if (i >= 0)
		{
			var id = NamedStyles[i].XfId;
			if (id < 0)
			{
				var styleXfId = NamedStyles[i].StyleXfId;
				var newStyle = CellStyleXfs[styleXfId].Copy();
				newStyle.XfId = styleXfId;
				id = CellXfs.FindIndexByID(newStyle.Id);
				if (id < 0)
				{
					id = CellXfs.Add(newStyle.Id, newStyle);
				}

				NamedStyles[i].XfId = id;
			}

			return id;
		}
		else
		{
			return 0;
		}
	}
}
