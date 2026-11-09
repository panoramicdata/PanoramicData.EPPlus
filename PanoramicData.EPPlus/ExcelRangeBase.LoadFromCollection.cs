using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Table;
using System.ComponentModel;
using System.Reflection;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Load a collection into a the worksheet starting from the top left row of the range.
	/// </summary>
	/// <typeparam name="T">The data type in the collection</typeparam>
	/// <param name="Collection">The collection to load</param>
	/// <returns>The filled range</returns>
	public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection) => LoadFromCollection<T>(Collection, false, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, null);
	/// <summary>
	/// Load a collection of T into the worksheet starting from the top left row of the range.
	/// Default option will load all public instance properties of T
	/// </summary>
	/// <typeparam name="T">The data type in the collection</typeparam>
	/// <param name="Collection">The collection to load</param>
	/// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
	/// <returns>The filled range</returns>
	public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders) => LoadFromCollection<T>(Collection, PrintHeaders, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, null);
	/// <summary>
	/// Load a collection of T into the worksheet starting from the top left row of the range.
	/// Default option will load all public instance properties of T
	/// </summary>
	/// <typeparam name="T">The datatype in the collection</typeparam>
	/// <param name="Collection">The collection to load</param>
	/// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
	/// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
	/// <returns>The filled range</returns>
	public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles TableStyle) => LoadFromCollection<T>(Collection, PrintHeaders, TableStyle, BindingFlags.Public | BindingFlags.Instance, null);
	/// <summary>
	/// Load a collection into the worksheet starting from the top left row of the range.
	/// </summary>
	/// <typeparam name="T">The datatype in the collection</typeparam>
	/// <param name="Collection">The collection to load</param>
	/// <param name="PrintHeaders">Print the property names on the first row. Any underscore in the property name will be converted to a space. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
	/// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
	/// <param name="memberFlags">Property flags to use</param>
	/// <param name="Members">The properties to output. Must be of type T</param>
	/// <returns>The filled range</returns>
	public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles TableStyle, BindingFlags memberFlags, MemberInfo[]? Members)
	{
		var type = typeof(T);
		var isSameType = true;
		if (Members == null)
		{
			Members = type.GetProperties(memberFlags);
		}
		else
		{
			if (Members.Length == 0)   //Fixes issue 15555
			{
				throw new ArgumentException("Parameter Members must have at least one property. Length is zero");
			}

			foreach (var t in Members)
			{
				if (t.DeclaringType != null && t.DeclaringType != type)
				{
					isSameType = false;
				}
				//Fixing inverted check for IsSubclassOf / Pullrequest from tomdam
				if (t.DeclaringType != null && t.DeclaringType != type && !TypeCompat.IsSubclassOf(type, t.DeclaringType) && !TypeCompat.IsSubclassOf(t.DeclaringType, type))
				{
					throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
				}
			}
		}

		// create buffer
		var values = new object[PrintHeaders ? Collection.Count() + 1 : Collection.Count(), Members.Length];

		int col = 0, row = 0;
		if (Members.Length > 0 && PrintHeaders)
		{
			foreach (var t in Members)
			{
				string header;
				if (t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() is DescriptionAttribute descriptionAttribute)
				{
					header = descriptionAttribute.Description;
				}
				else
				{
					if (t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() is DisplayNameAttribute displayNameAttribute)
					{
						header = displayNameAttribute.DisplayName;
					}
					else
					{
						header = t.Name.Replace('_', ' ');
					}
				}
				values[row, col++] = header;
			}

			row++;
		}

		if (!Collection.Any() && (Members.Length == 0 || PrintHeaders == false))
		{
			return null;
		}

		foreach (var item in Collection)
		{
			col = 0;
			if (item is string || item is decimal || item is DateTime || TypeCompat.IsPrimitive(item))
			{
				values[row, col++] = item;
			}
			else
			{
				foreach (var t in Members)
				{
					if (isSameType == false && item.GetType().GetMember(t.Name, memberFlags).Length == 0)
					{
						// The member does not exist on this inherited type
						col++;
					}
					else if (t is PropertyInfo info)
					{
						values[row, col++] = info.GetValue(item, null);
					}
					else if (t is FieldInfo info1)
					{
						values[row, col++] = info1.GetValue(item);
					}
					else if (t is MethodInfo info2)
					{
						values[row, col++] = info2.Invoke(item, null);
					}
				}
			}

			row++;
		}

		_worksheet.SetRangeValueInner(_fromRow, _fromCol, _fromRow + row - 1, _fromCol + col - 1, values);

		//Must have at least 1 row, if header is showen
		if (row == 1 && PrintHeaders)
		{
			row++;
		}

		var r = _worksheet.Cells[_fromRow, _fromCol, _fromRow + row - 1, _fromCol + col - 1];

		if (TableStyle != TableStyles.None)
		{
			var tbl = _worksheet.Tables.Add(r, "");
			tbl.ShowHeader = PrintHeaders;
			tbl.TableStyle = TableStyle;
		}

		return r;
	}
}
