using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;

namespace PanoramicData.EPPlus.Test.Utils;

[TestClass]
public class SqRefUtilityTests
{
	[TestMethod]
	public void SqRefUtility_ToSqRefAddress_ShouldThrowIfAddressIsNullOrEmpty() => Assert.ThrowsExactly<ArgumentNullException>(() => SqRefUtility.ToSqRefAddress(null));

	[TestMethod]
	public void SqRefUtility_ToSqRefAddress_ShouldRemoveCommas()
	{
		// Arrange
		var address = "A1, A2:A3";

		// Act
		var result = SqRefUtility.ToSqRefAddress(address);

		// Assert
		Assert.AreEqual("A1 A2:A3", result);
	}


	[TestMethod]
	public void SqRefUtility_ToSqRefAddress_ShouldRemoveCommasAndInsertSpaceIfNecesary()
	{
		// Arrange
		var address = "A1,A2:A3";

		// Act
		var result = SqRefUtility.ToSqRefAddress(address);

		// Assert
		Assert.AreEqual("A1 A2:A3", result);
	}

	[TestMethod]
	public void SqRefUtility_ToSqRefAddress_ShouldRemoveMultipleSpaces()
	{
		// Arrange
		var address = "A1,        A2:A3";

		// Act
		var result = SqRefUtility.ToSqRefAddress(address);

		// Assert
		Assert.AreEqual("A1 A2:A3", result);
	}

	[TestMethod]
	public void SqRefUtility_FromSqRefAddress_ShouldReplaceSpaceWithComma()
	{
		// Arrange
		var address = "A1 A2";

		// Act
		var result = SqRefUtility.FromSqRefAddress(address);

		// Assert
		Assert.AreEqual("A1,A2", result);
	}
}
