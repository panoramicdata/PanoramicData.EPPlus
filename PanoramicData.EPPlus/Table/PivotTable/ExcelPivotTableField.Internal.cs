using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public partial class ExcelPivotTableField
{
	#region Private & internal Methods
	internal static XmlElement AppendField(XmlNode rowsNode, int index, string fieldNodeText, string indexAttrText)
	{
		XmlElement prevField = null, newElement;
		foreach (XmlElement field in rowsNode.ChildNodes)
		{
			var x = field.GetAttribute(indexAttrText);
			if (int.TryParse(x, out var fieldIndex))
			{
				if (fieldIndex == index)    //Row already exists
				{
					return field;
				}
			}

			prevField = field;
		}

		newElement = rowsNode.OwnerDocument.CreateElement(fieldNodeText, ExcelPackage.schemaMain);
		newElement.SetAttribute(indexAttrText, index.ToString());
		rowsNode.InsertAfter(newElement, prevField);

		return newElement;
	}
	internal XmlHelperInstance _cacheFieldHelper;
	internal void SetCacheFieldNode(XmlNode cacheField)
	{
		_cacheFieldHelper = new XmlHelperInstance(NameSpaceManager, cacheField);
		var groupNode = cacheField.SelectSingleNode("d:fieldGroup", NameSpaceManager);
		if (groupNode != null)
		{
			var groupBy = groupNode.SelectSingleNode("d:rangePr/@groupBy", NameSpaceManager);
			if (groupBy == null)
			{
				_grouping = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, cacheField);
			}
			else
			{
				DateGrouping = Enum.Parse<eDateGroupBy>(groupBy.Value, true);
				_grouping = new ExcelPivotTableFieldDateGroup(NameSpaceManager, groupNode);
			}
		}
	}
	#endregion
}
