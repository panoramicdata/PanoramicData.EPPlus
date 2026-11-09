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
			int ret;
			for (var i = 0; i < columns.Length; i++)
			{
				var x1 = x.Items[columns[i]]._value;
				var y1 = y.Items[columns[i]]._value;
				var isNumX = ConvertUtil.IsNumeric(x1);
				var isNumY = ConvertUtil.IsNumeric(y1);
				if (isNumX && isNumY)   //Numeric Compare
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

					ret = d1 < d2 ? -1 : d1 > d2 ? 1 : 0;
				}
				else if (isNumX == false && isNumY == false)   //String Compare
				{
					var s1 = x1 == null ? "" : x1.ToString();
					var s2 = y1 == null ? "" : y1.ToString();
					ret = string.Compare(s1, s2, StringComparison.CurrentCulture);
				}
				else
				{
					ret = isNumX ? -1 : 1;
				}

				if (ret != 0)
				{
					return ret * (descending[i] ? -1 : 1);
				}
			}

			return 0;
		}
	}
}
