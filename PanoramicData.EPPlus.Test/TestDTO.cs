using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;

namespace PanoramicData.EPPlus.Test;

public class TestDTO
{
	[SuppressMessage("Major Code Smell", "S1104:Fields should not have public accessibility", Justification = "LoadFromCollection tests read public fields through reflection.")]
	[SuppressMessage("Minor Code Smell", "S2357:Fields should be private", Justification = "LoadFromCollection tests read public fields through reflection.")]
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
