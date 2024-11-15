﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;

namespace PanoramicData.EPPlus.Test.Utils;

[TestClass]
public class GuardingTests
{
	private class TestClass
	{

	}

	[TestMethod, ExpectedException(typeof(ArgumentNullException))]
	public void Require_IsNotNull_ShouldThrowIfArgumentIsNull()
	{
		TestClass obj = null;
		Require.Argument(obj).IsNotNull("test");
	}

	[TestMethod]
	public void Require_IsNotNull_ShouldNotThrowIfArgumentIsAnInstance()
	{
		var obj = new TestClass();
		Require.Argument(obj).IsNotNull("test");
	}

	[TestMethod, ExpectedException(typeof(ArgumentNullException))]
	public void Require_IsNotNullOrEmpty_ShouldThrowIfStringIsNull()
	{
		string arg = null;
		Require.Argument(arg).IsNotNullOrEmpty("test");
	}

	[TestMethod]
	public void Require_IsNotNullOrEmpty_ShouldNotThrowIfStringIsNotNullOrEmpty()
	{
		var arg = "test";
		Require.Argument(arg).IsNotNullOrEmpty("test");
	}

	[TestMethod, ExpectedException(typeof(ArgumentOutOfRangeException))]
	public void Require_IsInRange_ShouldThrowIfArgumentIsOutOfRange()
	{
		var arg = 3;
		Require.Argument(arg).IsInRange(5, 7, "test");
	}

	[TestMethod]
	public void Require_IsInRange_ShouldNotThrowIfArgumentIsInRange()
	{
		var arg = 6;
		Require.Argument(arg).IsInRange(5, 7, "test");
	}
}
