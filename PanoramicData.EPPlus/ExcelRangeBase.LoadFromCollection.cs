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
		var isSameType = ResolveCollectionMembers(type, memberFlags, ref Members);

		// create buffer
		var values = new object[PrintHeaders ? Collection.Count() + 1 : Collection.Count(), Members.Length];

		int col = 0, row = 0;
		if (Members.Length > 0 && PrintHeaders)
		{
			col = WriteCollectionHeaders(values, Members);
			row++;
		}

		if (IsEmptyCollectionLoad(Collection, Members.Length, PrintHeaders))
		{
			return null;
		}

		foreach (var item in Collection)
		{
			col = WriteCollectionItem(values, row, item, Members, isSameType, memberFlags);
			row++;
		}

		return CreateCollectionRange(values, row, col, PrintHeaders, TableStyle);
	}

	private static bool ResolveCollectionMembers(Type type, BindingFlags memberFlags, ref MemberInfo[]? Members)
	{
		if (Members == null)
		{
			Members = [.. type.GetProperties(memberFlags)];
			return true;
		}

		if (Members.Length == 0)   //Fixes issue 15555
		{
			throw new ArgumentException("Parameter Members must have at least one property. Length is zero");
		}

		var isSameType = true;
		foreach (var t in Members)
		{
			if (IsDeclaredOnOtherType(t, type))
			{
				isSameType = false;
			}
			//Fixing inverted check for IsSubclassOf / Pullrequest from tomdam
			if (IsDeclaredOnOtherType(t, type) && IsUnrelatedType(type, t.DeclaringType))
			{
				throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
			}
		}

		return isSameType;
	}

	private static bool IsDeclaredOnOtherType(MemberInfo t, Type type) => t.DeclaringType != null && t.DeclaringType != type;

	private static bool IsUnrelatedType(Type type, Type declaringType)
		=> !TypeCompat.IsSubclassOf(type, declaringType) && !TypeCompat.IsSubclassOf(declaringType, type);

	private static int WriteCollectionHeaders(object[,] values, MemberInfo[] Members)
	{
		var col = 0;
		foreach (var t in Members)
		{
			values[0, col++] = GetMemberHeader(t);
		}

		return col;
	}

	private static string GetMemberHeader(MemberInfo t)
	{
		if (t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() is DescriptionAttribute descriptionAttribute)
		{
			return descriptionAttribute.Description;
		}
		else if (t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() is DisplayNameAttribute displayNameAttribute)
		{
			return displayNameAttribute.DisplayName;
		}
		else
		{
			return t.Name.Replace('_', ' ');
		}
	}

	private static bool IsEmptyCollectionLoad<T>(IEnumerable<T> Collection, int memberCount, bool PrintHeaders)
		=> !Collection.Any() && (memberCount == 0 || PrintHeaders == false);

	private static int WriteCollectionItem(object[,] values, int row, object item, MemberInfo[] Members, bool isSameType, BindingFlags memberFlags)
	{
		var col = 0;
		if (IsScalarItem(item))
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
				else if (TryGetMemberValue(t, item, out var value))
				{
					values[row, col++] = value;
				}
			}
		}

		return col;
	}

	private static bool IsScalarItem(object item)
		=> item is string || item is decimal || item is DateTime || TypeCompat.IsPrimitive(item);

	private static bool TryGetMemberValue(MemberInfo t, object item, out object value)
	{
		if (t is PropertyInfo info)
		{
			value = info.GetValue(item, null);
			return true;
		}
		else if (t is FieldInfo info1)
		{
			value = info1.GetValue(item);
			return true;
		}
		else if (t is MethodInfo info2)
		{
			value = info2.Invoke(item, null);
			return true;
		}

		value = null;
		return false;
	}

	private ExcelRangeBase CreateCollectionRange(object[,] values, int row, int col, bool PrintHeaders, TableStyles TableStyle)
	{
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
