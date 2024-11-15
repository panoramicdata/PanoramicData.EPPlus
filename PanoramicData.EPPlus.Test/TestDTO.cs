﻿using System;
using System.ComponentModel;

namespace PanoramicData.EPPlus.Test;

public class TestDTO
{
	public string NameVar;

	public int Id { get; set; }
	[DisplayName("Name from DisplayNameAttribute")]
	public string Name { get; set; }
	public TestDTO? dto { get; set; }
	public DateTime Date { get; set; }
	public bool Boolean { get; set; }

	public string GetNameID() => Id + "," + Name;
}
public class InheritTestDTO : TestDTO
{
	public string InheritedProp { get; set; }
}
