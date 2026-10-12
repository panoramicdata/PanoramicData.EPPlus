using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Globalization;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// The suite asserts on culture-dependent text (dates, built-in number formats), so it runs under
/// one fixed culture rather than whatever the machine is set to. Without this, a Linux runner
/// (usually the invariant culture) and a developer's machine (any locale) would disagree.
/// The culture is created without user overrides so Windows regional settings do not leak in.
/// </summary>
[TestClass]
public static class TestAssemblySetup
{
	[AssemblyInitialize]
	public static void SetCulture(TestContext _)
	{
		var culture = new CultureInfo("en-US", false);
		CultureInfo.DefaultThreadCurrentCulture = culture;
		CultureInfo.DefaultThreadCurrentUICulture = culture;
		CultureInfo.CurrentCulture = culture;
		CultureInfo.CurrentUICulture = culture;
	}
}
