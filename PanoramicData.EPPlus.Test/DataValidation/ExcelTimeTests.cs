﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace PanoramicData.EPPlus.Test.DataValidation;

[TestClass]
public class ExcelTimeTests
{
	private ExcelTime _time;
	private readonly decimal SecondsPerHour = 3600;
	// private readonly decimal HoursPerDay = 24;
	private readonly decimal SecondsPerDay = 3600 * 24;

	private static decimal Round(decimal value) => Math.Round(value, ExcelTime.NumberOfDecimals);

	[TestInitialize]
	public void Setup() => _time = new ExcelTime();

	[TestCleanup]
	public void Cleanup() => _time = null;

	[TestMethod, ExpectedException(typeof(ArgumentException))]
	public void ExcelTimeTests_ConstructorWithValue_ShouldThrowIfValueIsLessThan0() => new ExcelTime(-1);

	[TestMethod, ExpectedException(typeof(ArgumentException))]
	public void ExcelTimeTests_ConstructorWithValue_ShouldThrowIfValueIsEqualToOrGreaterThan1() => new ExcelTime(1);

	[TestMethod, ExpectedException(typeof(InvalidOperationException))]
	public void ExcelTimeTests_Hour_ShouldThrowIfNegativeValue() => _time.Hour = -1;

	[TestMethod, ExpectedException(typeof(InvalidOperationException))]
	public void ExcelTimeTests_Minute_ShouldThrowIfNegativeValue() => _time.Minute = -1;

	[TestMethod, ExpectedException(typeof(InvalidOperationException))]
	public void ExcelTimeTests_Minute_ShouldThrowIValueIsGreaterThan59() => _time.Minute = 60;

	[TestMethod, ExpectedException(typeof(InvalidOperationException))]
	public void ExcelTimeTests_Second_ShouldThrowIfNegativeValue() => _time.Second = -1;

	[TestMethod, ExpectedException(typeof(InvalidOperationException))]
	public void ExcelTimeTests_Second_ShouldThrowIValueIsGreaterThan59() => _time.Second = 60;

	[TestMethod]
	public void ExcelTimeTests_ToExcelTime_HourIsSet()
	{
		// Act
		_time.Hour = 1;

		// Assert
		Assert.AreEqual(Round(SecondsPerHour / SecondsPerDay), _time.ToExcelTime());
	}

	[TestMethod]
	public void ExcelTimeTests_ToExcelTime_MinuteIsSet()
	{
		// Arrange
		var expected = SecondsPerHour + 20M * 60M;
		// Act
		_time.Hour = 1;
		_time.Minute = 20;

		// Assert
		Assert.AreEqual(Round(expected / SecondsPerDay), _time.ToExcelTime());
	}

	[TestMethod]
	public void ExcelTimeTests_ToExcelTime_SecondIsSet()
	{
		// Arrange
		var expected = SecondsPerHour + 20M * 60M + 10M;
		// Act
		_time.Hour = 1;
		_time.Minute = 20;
		_time.Second = 10;

		// Assert
		Assert.AreEqual(Round(expected / SecondsPerDay), _time.ToExcelTime());
	}

	[TestMethod]
	public void ExcelTimeTests_ConstructorWithValue_ShouldSetHour()
	{
		// Arrange
		var value = 3660M / SecondsPerDay;

		// Act
		var time = new ExcelTime(value);

		// Assert
		Assert.AreEqual(1, time.Hour);
	}

	[TestMethod]
	public void ExcelTimeTests_ConstructorWithValue_ShouldSetMinute()
	{
		// Arrange
		var value = 3660M / SecondsPerDay;

		// Act
		var time = new ExcelTime(value);

		// Assert
		Assert.AreEqual(1, time.Minute);
	}

	[TestMethod]
	public void ExcelTimeTests_ConstructorWithValue_ShouldSetSecond()
	{
		// Arrange
		var value = 3662M / SecondsPerDay;

		// Act
		var time = new ExcelTime(value);

		// Assert
		Assert.AreEqual(2, time.Second);
	}
}
