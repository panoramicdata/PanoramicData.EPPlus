using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public partial class ExcelPivotTableField
{
	#region Grouping
	internal ExcelPivotTableFieldDateGroup SetDateGroup(eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
	{
		ExcelPivotTableFieldDateGroup group;
		group = new ExcelPivotTableFieldDateGroup(NameSpaceManager, _cacheFieldHelper.TopNode);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsDate", true);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsNonDate", false);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);

		group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"{1}\" /><groupItems /></fieldGroup>", BaseIndex, GroupBy.ToString().ToLower(CultureInfo.InvariantCulture));

		if (StartDate.Year < 1900)
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", "1900-01-01T00:00:00");
		}
		else
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", StartDate.ToString("s", CultureInfo.InvariantCulture));
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoStart", "0");
		}

		if (EndDate == DateTime.MaxValue)
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", "9999-12-31T00:00:00");
		}
		else
		{
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", EndDate.ToString("s", CultureInfo.InvariantCulture));
			_cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoEnd", "0");
		}

		var items = AddDateGroupItems(group, GroupBy, StartDate, EndDate, interval);
		AddFieldItems(items);

		_grouping = group;
		return group;
	}
	internal ExcelPivotTableFieldNumericGroup SetNumericGroup(double start, double end, double interval)
	{
		ExcelPivotTableFieldNumericGroup group;
		group = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, _cacheFieldHelper.TopNode);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsNumber", true);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsInteger", true);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);
		_cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsString", false);

		group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"{1}\" endNum=\"{2}\" groupInterval=\"{3}\"/><groupItems /></fieldGroup>", BaseIndex, start.ToString(CultureInfo.InvariantCulture), end.ToString(CultureInfo.InvariantCulture), interval.ToString(CultureInfo.InvariantCulture));
		var items = AddNumericGroupItems(group, start, end, interval);
		AddFieldItems(items);

		_grouping = group;
		return group;
	}

	private static int AddNumericGroupItems(ExcelPivotTableFieldNumericGroup group, double start, double end, double interval)
	{
		if (interval < 0)
		{
			throw (new Exception("The interval must be a positiv"));
		}

		if (start > end)
		{
			throw (new Exception("Then End number must be larger than the Start number"));
		}

		var groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
		var items = 2;
		//First date
		var index = start;
		var nextIndex = start + interval;
		AddGroupItem(groupItems, "<" + start.ToString(CultureInfo.InvariantCulture));

		while (index < end)
		{
			AddGroupItem(groupItems, string.Format("{0}-{1}", index.ToString(CultureInfo.InvariantCulture), nextIndex.ToString(CultureInfo.InvariantCulture)));
			index = nextIndex;
			nextIndex += interval;
			items++;
		}

		AddGroupItem(groupItems, ">" + nextIndex.ToString(CultureInfo.InvariantCulture));
		return items;
	}

	private void AddFieldItems(int items)
	{
		XmlElement prevNode = null;
		var itemsNode = TopNode.SelectSingleNode("d:items", NameSpaceManager) as XmlElement;
		for (var x = 0; x < items; x++)
		{
			var itemNode = itemsNode.OwnerDocument.CreateElement("item", ExcelPackage.schemaMain);
			itemNode.SetAttribute("x", x.ToString());
			if (prevNode == null)
			{
				itemsNode.PrependChild(itemNode);
			}
			else
			{
				itemsNode.InsertAfter(itemNode, prevNode);
			}

			prevNode = itemNode;
		}

		itemsNode.SetAttribute("count", (items + 1).ToString());
	}

	private int AddDateGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
	{
		var groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
		var items = 2;
		//First date
		AddGroupItem(groupItems, "<" + StartDate.ToString("s", CultureInfo.InvariantCulture)[..10]);

		switch (GroupBy)
		{
			case eDateGroupBy.Seconds:
			case eDateGroupBy.Minutes:
				AddTimeSerie(60, groupItems);
				items += 60;
				break;
			case eDateGroupBy.Hours:
				AddTimeSerie(24, groupItems);
				items += 24;
				break;
			case eDateGroupBy.Days:
				items = AddDayItems(groupItems, StartDate, EndDate, interval, items);
				break;
			case eDateGroupBy.Months:
				AddGroupItems(groupItems, _monthItems);
				items += 12;
				break;
			case eDateGroupBy.Quarters:
				AddGroupItems(groupItems, _quarterItems);
				items += 4;
				break;
			case eDateGroupBy.Years:
				items += AddYearItems(groupItems, StartDate, EndDate);
				break;
			default:
				throw (new Exception("unsupported grouping"));
		}

		//Lastdate
		AddGroupItem(groupItems, ">" + EndDate.ToString("s", CultureInfo.InvariantCulture)[..10]);
		return items;
	}

	static readonly string[] _monthItems = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
	static readonly string[] _quarterItems = ["Qtr1", "Qtr2", "Qtr3", "Qtr4"];

	private static void AddGroupItems(XmlElement groupItems, string[] values)
	{
		foreach (var value in values)
		{
			AddGroupItem(groupItems, value);
		}
	}

	// Returns the new item count: an interval of one lists every day of a leap year, otherwise each interval in range.
	private static int AddDayItems(XmlElement groupItems, DateTime startDate, DateTime endDate, int interval, int items)
	{
		if (interval == 1)
		{
			DateTime dt = new(2008, 1, 1); //pick a year with 366 days
			while (dt.Year == 2008)
			{
				AddGroupItem(groupItems, dt.ToString("dd-MMM"));
				dt = dt.AddDays(1);
			}

			return items + 366;
		}

		var day = startDate;
		var count = 0;
		while (day < endDate)
		{
			AddGroupItem(groupItems, day.ToString("dd-MMM"));
			day = day.AddDays(interval);
			count++;
		}

		return count;
	}

	private static int AddYearItems(XmlElement groupItems, DateTime startDate, DateTime endDate)
	{
		if (startDate.Year < 1900 || endDate == DateTime.MaxValue)
		{
			return 0;
		}

		for (var year = startDate.Year; year <= endDate.Year; year++)
		{
			AddGroupItem(groupItems, year.ToString());
		}

		return endDate.Year - startDate.Year + 1;
	}

	private static void AddTimeSerie(int count, XmlElement groupItems)
	{
		for (var i = 0; i < count; i++)
		{
			AddGroupItem(groupItems, string.Format("{0:00}", i));
		}
	}

	private static void AddGroupItem(XmlElement groupItems, string value)
	{
		var s = groupItems.OwnerDocument.CreateElement("s", ExcelPackage.schemaMain);
		s.SetAttribute("v", value);
		groupItems.AppendChild(s);
	}
	#endregion
}
