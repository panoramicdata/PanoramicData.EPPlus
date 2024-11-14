/* Copyright (C) 2011  Jan Källman
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 *
 * Author Change                      Date
 *******************************************************************************
 * Mats Alm Added		                2016-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing;

/// <summary>
/// Represents the errortypes in excel
/// </summary>
public enum eErrorType
{
	/// <summary>
	/// Division by zero
	/// </summary>
	Div0,
	/// <summary>
	/// Not applicable
	/// </summary>
	NA,
	/// <summary>
	/// Name error
	/// </summary>
	Name,
	/// <summary>
	/// Null error
	/// </summary>
	Null,
	/// <summary>
	/// Num error
	/// </summary>
	Num,
	/// <summary>
	/// Reference error
	/// </summary>
	Ref,
	/// <summary>
	/// Value error
	/// </summary>
	Value
}

/// <summary>
/// Represents an Excel error.
/// </summary>
/// <seealso cref="eErrorType"/>
public class ExcelErrorValue
{
	/// <summary>
	/// Handles the convertion between <see cref="eErrorType"/> and the string values
	/// used by Excel.
	/// </summary>
	public static class Values
	{
		public const string Div0 = "#DIV/0!";
		public const string NA = "#N/A";
		public const string Name = "#NAME?";
		public const string Null = "#NULL!";
		public const string Num = "#NUM!";
		public const string Ref = "#REF!";
		public const string Value = "#VALUE!";

		private static readonly Dictionary<string, eErrorType> _values = new()
			{
				{Div0, eErrorType.Div0},
				{NA, eErrorType.NA},
				{Name, eErrorType.Name},
				{Null, eErrorType.Null},
				{Num, eErrorType.Num},
				{Ref, eErrorType.Ref},
				{Value, eErrorType.Value}
			};

		/// <summary>
		/// Returns true if the supplied <paramref name="candidate"/> is an excel error.
		/// </summary>
		/// <param name="candidate"></param>
		/// <returns></returns>
		public static bool IsErrorValue(object candidate)
		{
			if (candidate is null or not ExcelErrorValue) return false;
			var candidateString = candidate.ToString();
			return !string.IsNullOrEmpty(candidateString) && _values.ContainsKey(candidateString);
		}

		/// <summary>
		/// Returns true if the supplied <paramref name="candidate"/> is an excel error.
		/// </summary>
		/// <param name="candidate"></param>
		/// <returns></returns>
		public static bool StringIsErrorValue(string candidate) => !string.IsNullOrEmpty(candidate) && _values.ContainsKey(candidate);

		/// <summary>
		/// Converts a string to an <see cref="eErrorType"/>
		/// </summary>
		/// <param name="val"></param>
		/// <returns></returns>
		/// <exception cref="ArgumentException">Thrown if the supplied value is not an Excel error</exception>
		public static eErrorType ToErrorType(string val) => string.IsNullOrEmpty(val) || !_values.TryGetValue(val, out var value)
				? throw new ArgumentException("Invalid error code " + (val ?? "<empty>"))
				: value;
	}

	internal static ExcelErrorValue Create(eErrorType errorType) => new(errorType);

	internal static ExcelErrorValue Parse(string val)
	{
		if (Values.StringIsErrorValue(val))
		{
			return new ExcelErrorValue(Values.ToErrorType(val));
		}

		if (string.IsNullOrEmpty(val)) throw new ArgumentNullException(nameof(val));
		throw new ArgumentException("Not a valid error value: " + val);
	}

	private ExcelErrorValue(eErrorType type)
	{
		Type = type;
	}

	/// <summary>
	/// The error type
	/// </summary>
	public eErrorType Type { get; private set; }

	/// <summary>
	/// Returns the string representation of the error type
	/// </summary>
	/// <returns></returns>
	public override string ToString() => Type switch
	{
		eErrorType.Div0 => Values.Div0,
		eErrorType.NA => Values.NA,
		eErrorType.Name => Values.Name,
		eErrorType.Null => Values.Null,
		eErrorType.Num => Values.Num,
		eErrorType.Ref => Values.Ref,
		eErrorType.Value => Values.Value,
		_ => throw (new ArgumentException("Invalid errortype")),
	};
	public static ExcelErrorValue operator +(object v1, ExcelErrorValue v2)
	{
		return v2;
	}
	public static ExcelErrorValue operator +(ExcelErrorValue v1, ExcelErrorValue v2)
	{
		return v1;
	}

	public override int GetHashCode() => base.GetHashCode();

	public override bool Equals(object obj) => obj is ExcelErrorValue && ((ExcelErrorValue)obj).ToString() == ToString();
}
