using OfficeOpenXml.Compatibility;
using System.ComponentModel;
using System.Reflection;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
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
}
