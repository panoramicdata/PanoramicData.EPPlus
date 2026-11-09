using OfficeOpenXml.Utils;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	[SuppressMessage("Minor Code Smell", "S3898:Value types should implement IEquatable<T>", Justification = "Private sort buffer item; never compared for equality.")]
	private struct SortItem<T>
	{
		internal int Row { get; set; }
		internal T[] Items { get; set; }
	}
	private class Comp : IComparer<SortItem<ExcelCoreValue>>
	{
		public int[] columns;
		public bool[] descending;
		public CultureInfo cultureInfo = CultureInfo.CurrentCulture;
		public CompareOptions compareOptions = CompareOptions.None;
		public int Compare(SortItem<ExcelCoreValue> x, SortItem<ExcelCoreValue> y)
		{
			for (var i = 0; i < columns.Length; i++)
			{
				var ret = CompareValues(x.Items[columns[i]]._value, y.Items[columns[i]]._value);
				if (ret != 0)
				{
					return ret * (descending[i] ? -1 : 1);
				}
			}

			return 0;
		}

		private static int CompareValues(object x1, object y1)
		{
			var isNumX = ConvertUtil.IsNumeric(x1);
			var isNumY = ConvertUtil.IsNumeric(y1);
			if (isNumX && isNumY)   //Numeric Compare
			{
				return CompareNumbers(x1, y1);
			}
			else if (isNumX == false && isNumY == false)   //String Compare
			{
				var s1 = AsSortText(x1);
				var s2 = AsSortText(y1);
				return string.Compare(s1, s2, StringComparison.CurrentCulture);
			}
			else
			{
				return isNumX ? -1 : 1;
			}
		}

		private static int CompareNumbers(object x1, object y1)
		{
			var d1 = ConvertUtil.GetValueDouble(x1);
			var d2 = ConvertUtil.GetValueDouble(y1);
			if (double.IsNaN(d1))
			{
				d1 = double.MaxValue;
			}

			if (double.IsNaN(d2))
			{
				d2 = double.MaxValue;
			}

			return d1 < d2 ? -1 : d1 > d2 ? 1 : 0;
		}

		private static string AsSortText(object value) => value == null ? "" : value.ToString();
	}
}
