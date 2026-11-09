using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace PanoramicData.EPPlus.Test;

[TestClass]
public class LoadFromCollectionTests
{
	internal abstract class BaseClass
	{
		public string Id { get; set; }
		public string Name { get; set; }
	}

	internal class Implementation : BaseClass
	{
		public int Number { get; set; }
	}

	internal class Aclass
	{
		public string Id { get; set; }
		public string Name { get; set; }
		public int Number { get; set; }
	}

	private static List<BaseClass> CreateImplementations() =>
	[
		new Implementation { Id = "123", Name = "Item 1", Number = 3 }
	];

	private static object LoadIntoC1<T>(IEnumerable<T> items)
	{
		using var pck = new ExcelPackage(new MemoryStream());
		var sheet = pck.Workbook.Worksheets.Add("sheet");
		sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);
		return sheet.Cells["C1"].Value;
	}

	[TestMethod]
	public void ShouldUseAclassProperties()
	{
		var items = new List<Aclass>
		{
			new(){ Id = "123", Name = "Item 1", Number = 3}
		};
		Assert.AreEqual("Id", LoadIntoC1(items));
	}

	[TestMethod]
	public void ShouldUseBaseClassProperties() => Assert.AreEqual("Id", LoadIntoC1(CreateImplementations()));

	[TestMethod]
	public void ShouldUseAnonymousProperties()
	{
		var items = CreateImplementations().Select(x => new { x.Id, x.Name }).ToList();
		Assert.AreEqual("Id", LoadIntoC1(items));
	}

	[TestMethod]
	public void ShouldThrowInvalidCastExceptionIf()
	{
		var items = CreateImplementations().Select(x => new { x.Id, x.Name }).ToList();
		using var pck = new ExcelPackage(new MemoryStream());
		var sheet = pck.Workbook.Worksheets.Add("sheet");
		Assert.ThrowsExactly<InvalidCastException>(() => sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1, BindingFlags.Public | BindingFlags.Instance, typeof(string).GetMembers()));
	}
}