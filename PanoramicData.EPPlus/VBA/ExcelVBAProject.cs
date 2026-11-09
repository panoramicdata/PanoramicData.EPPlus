/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * If you want to understand this code have a look at the Office VBA File Format Structure Specification (MS-OVBA.PDF) or
 * http://msdn.microsoft.com/en-us/library/cc313094(v=office.12).aspx
 * 
 * * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		26-MAR-2012
 *******************************************************************************/
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.VBA;

/// <summary>
/// Represents the VBA project part of the package
/// </summary>
public partial class ExcelVbaProject
{
	const string schemaRelVba = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
	internal const string PartUri = @"/xl/vbaProject.bin";
	#region Classes & Enums
	/// <summary>
	/// Type of system where the VBA project was created.
	/// </summary>
	public enum eSyskind
	{
		Win16 = 0,
		Win32 = 1,
		Macintosh = 2,
		Win64 = 3
	}

	#endregion
	internal ExcelVbaProject(ExcelWorkbook wb)
	{
		_wb = wb;
		_pck = _wb._package.Package;
		References = [];
		Modules = new ExcelVbaModuleCollection(this);
		var rel = _wb.Part.GetRelationshipsByType(schemaRelVba).FirstOrDefault();
		if (rel != null)
		{
			Uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
			Part = _pck.GetPart(Uri);
			GetProject();
		}
		else
		{
			Lcid = 0;
			Part = null;
		}
	}
	internal ExcelWorkbook _wb;
	internal Packaging.ZipPackage _pck;
	#region Dir Stream Properties
	/// <summary>
	/// System kind. Default Win32.
	/// </summary>
	public eSyskind SystemKind { get; set; }
	/// <summary>
	/// Name of the project
	/// </summary>
	public string Name { get; set; }
	/// <summary>
	/// A description of the project
	/// </summary>
	public string Description { get; set; }
	/// <summary>
	/// A helpfile
	/// </summary>
	public string HelpFile1 { get; set; }
	/// <summary>
	/// Secondary helpfile
	/// </summary>
	public string HelpFile2 { get; set; }
	/// <summary>
	/// Context if refering the helpfile
	/// </summary>
	public int HelpContextID { get; set; }
	/// <summary>
	/// Conditional compilation constants
	/// </summary>
	public string Constants { get; set; }
	/// <summary>
	/// Codepage for encoding. Default is current regional setting.
	/// </summary>
	public int CodePage { get; internal set; }
	internal int LibFlags { get; set; }
	internal int MajorVersion { get; set; }
	internal int MinorVersion { get; set; }
	internal int Lcid { get; set; }
	internal int LcidInvoke { get; set; }
	internal string ProjectID { get; set; }
	internal string ProjectStreamText { get; set; }
	/// <summary>
	/// Project references
	/// </summary>        
	public ExcelVbaReferenceCollection References { get; set; }
	/// <summary>
	/// Code Modules (Modules, classes, designer code)
	/// </summary>
	public ExcelVbaModuleCollection Modules { get; set; }
	ExcelVbaSignature _signature;
	/// <summary>
	/// The digital signature
	/// </summary>
	public ExcelVbaSignature Signature
	{
		get
		{
			_signature ??= new ExcelVbaSignature(Part);

			return _signature;
		}
	}
	ExcelVbaProtection _protection;
	/// <summary>
	/// VBA protection 
	/// </summary>
	public ExcelVbaProtection Protection
	{
		get
		{
			_protection ??= new ExcelVbaProtection(this);

			return _protection;
		}
	}
	#endregion

	private string GetString(BinaryReader br, uint size) => GetString(br, size, Encoding.GetEncoding(CodePage));
	private static string GetString(BinaryReader br, uint size, Encoding enc)
	{
		if (size > 0)
		{
			var byteTemp = br.ReadBytes((int)size);
			return enc.GetString(byteTemp);
		}
		else
		{
			return "";
		}
	}
	private string GetUnicodeString(BinaryReader br, uint size)
	{
		var s = GetString(br, size);
		_ = br.ReadUInt16();
		var sizeUC = br.ReadUInt32();
		var sUC = GetString(br, sizeUC, Encoding.Unicode);
		return sUC.Length == 0 ? s : sUC;
	}
	internal CompoundDocument Document { get; set; }
	internal Packaging.ZipPackagePart Part { get; set; }
	internal Uri Uri { get; private set; }
	/// <summary>
	/// Create a new VBA Project
	/// </summary>
	internal void Create()
	{
		if (Lcid > 0)
		{
			throw (new InvalidOperationException("Package already contains a VBAProject"));
		}

		ProjectID = "{5DD90D76-4904-47A2-AF0D-D69B4673604E}";
		Name = "VBAProject";
		SystemKind = eSyskind.Win32;            //Default
		Lcid = 1033;                            //English - United States
		LcidInvoke = 1033;                      //English - United States
		CodePage = Encoding.GetEncoding(0).CodePage;    //Switched from Default to make it work in Core
		MajorVersion = 1361024421;
		MinorVersion = 6;
		HelpContextID = 0;
		Modules.Add(new ExcelVBAModule(_wb.CodeNameChange) { Name = "ThisWorkbook", Code = "", Attributes = GetDocumentAttributes("ThisWorkbook", "0{00020819-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
		foreach (var sheet in _wb.Worksheets)
		{
			var name = GetModuleNameFromWorksheet(sheet);
			if (!Modules.Exists(name))
			{
				Modules.Add(new ExcelVBAModule(sheet.CodeNameChange) { Name = name, Code = "", Attributes = GetDocumentAttributes(sheet.Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
			}
		}

		_protection = new ExcelVbaProtection(this) { UserProtected = false, HostProtected = false, VbeProtected = false, VisibilityState = true };
	}

	internal string GetModuleNameFromWorksheet(ExcelWorksheet sheet)
	{
		var name = sheet.Name;
		name = name[..(name.Length < 31 ? name.Length : 31)];  //Maximum 31 charachters
		if (Modules[name] != null || !Regex.IsMatch(name, "^[a-zA-Z][a-zA-Z0-9_ ]*$")) //Check for valid chars, if not valid, set to sheetX.
		{
			var i = sheet.PositionID;
			name = "Sheet" + i.ToString();
			while (Modules[name] != null)
			{
				name = "Sheet" + (++i).ToString();
			}
		}

		return name;
	}
	internal static ExcelVbaModuleAttributesCollection GetDocumentAttributes(string name, string clsid)
	{
		var attr = new ExcelVbaModuleAttributesCollection();
		attr._list.Add(new ExcelVbaModuleAttribute { Name = "VB_Name", Value = name, DataType = eAttributeDataType.String });
		attr._list.Add(new ExcelVbaModuleAttribute { Name = "VB_Base", Value = clsid, DataType = eAttributeDataType.String });
		attr._list.Add(new ExcelVbaModuleAttribute { Name = "VB_GlobalNameSpace", Value = "False", DataType = eAttributeDataType.NonString });
		attr._list.Add(new ExcelVbaModuleAttribute { Name = "VB_Creatable", Value = "False", DataType = eAttributeDataType.NonString });
		attr._list.Add(new ExcelVbaModuleAttribute { Name = "VB_PredeclaredId", Value = "True", DataType = eAttributeDataType.NonString });
		attr._list.Add(new ExcelVbaModuleAttribute { Name = "VB_Exposed", Value = "False", DataType = eAttributeDataType.NonString });
		attr._list.Add(new ExcelVbaModuleAttribute { Name = "VB_TemplateDerived", Value = "False", DataType = eAttributeDataType.NonString });
		attr._list.Add(new ExcelVbaModuleAttribute { Name = "VB_Customizable", Value = "True", DataType = eAttributeDataType.NonString });

		return attr;
	}
	/// <summary>
	/// Remove the project from the package
	/// </summary>
	public void Remove()
	{
		if (Part == null)
		{
			return;
		}

		foreach (var rel in Part.GetRelationships())
		{
			_pck.DeleteRelationship(rel.Id);
		}

		if (_pck.PartExists(Uri))
		{
			_pck.DeletePart(Uri);
		}

		Part = null;
		Modules.Clear();
		References.Clear();
		Lcid = 0;
		LcidInvoke = 0;
		CodePage = 0;
		MajorVersion = 0;
		MinorVersion = 0;
		HelpContextID = 0;
	}
	public override string ToString() => Name;
}
