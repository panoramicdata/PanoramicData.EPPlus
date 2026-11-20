using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System.Text;

namespace OfficeOpenXml.VBA;

public partial class ExcelVbaProject
{
	#region Save Project
	internal void Save()
	{
		if (Validate())
		{
			CompoundDocument doc = new()
			{
				Storage = new CompoundDocument.StoragePart()
			};
			var store = new CompoundDocument.StoragePart();
			doc.Storage.SubStorage.Add("VBA", store);

			store.DataStreams.Add("_VBA_PROJECT", CreateVBAProjectStream());
			store.DataStreams.Add("dir", CreateDirStream());
			foreach (var module in Modules)
			{
				store.DataStreams.Add(module.Name, VBACompression.CompressPart(Encoding.GetEncoding(CodePage).GetBytes(module.Attributes.GetAttributeText() + module.Code)));
			}

			//Copy streams from the template, if used.
			if (Document != null)
			{
				foreach (var ss in Document.Storage.SubStorage)
				{
					if (ss.Key != "VBA")
					{
						doc.Storage.SubStorage.Add(ss.Key, ss.Value);
					}
				}

				foreach (var s in Document.Storage.DataStreams)
				{
					if (s.Key is not "dir" and not "PROJECT" and not "PROJECTwm")
					{
						doc.Storage.DataStreams.Add(s.Key, s.Value);
					}
				}
			}

			doc.Storage.DataStreams.Add("PROJECT", CreateProjectStream());
			doc.Storage.DataStreams.Add("PROJECTwm", CreateProjectwmStream());

			if (Part == null)
			{
				Uri = new Uri(PartUri, UriKind.Relative);
				Part = _pck.CreatePart(Uri, ExcelPackage.schemaVBA);
				_wb.Part.CreateRelationship(Uri, Packaging.TargetMode.Internal, schemaRelVba);
			}

			var st = Part.GetStream(FileMode.Create);
			doc.Save(st);
			st.Flush();
			//Save the digital signture
			Signature.Save(this);
		}
	}

	private bool Validate()
	{
		Description = Description ?? "";
		HelpFile1 = HelpFile1 ?? "";
		HelpFile2 = HelpFile2 ?? "";
		Constants = Constants ?? "";
		return true;
	}

	/// <summary>
	/// MS-OVBA 2.3.4.1
	/// </summary>
	/// <returns></returns>
	private static byte[] CreateVBAProjectStream()
	{
		BinaryWriter bw = new(new MemoryStream());
		bw.Write((ushort)0x61CC); //Reserved1
		bw.Write((ushort)0xFFFF); //Version
		bw.Write((byte)0x0); //Reserved3
		bw.Write((ushort)0x0); //Reserved4
		return ((MemoryStream)bw.BaseStream).ToArray();
	}
	/// <summary>
	/// MS-OVBA 2.3.4.1
	/// </summary>
	/// <returns></returns>
	private byte[] CreateDirStream()
	{
		BinaryWriter bw = new(new MemoryStream());
		WriteProjectInformation(bw);

		/****** PROJECTREFERENCES Record ******/
		foreach (var reference in References)
		{
			WriteNameReference(bw, reference);

			if (reference.ReferenceRecordID == 0x2F)
			{
				WriteControlReference(bw, reference);
			}
			else if (reference.ReferenceRecordID == 0x33)
			{
				WriteOrginalReference(bw, reference);
			}
			else if (reference.ReferenceRecordID == 0x0D)
			{
				WriteRegisteredReference(bw, reference);
			}
			else if (reference.ReferenceRecordID == 0x0E)
			{
				WriteProjectReference(bw, reference);
			}
		}

		bw.Write((ushort)0x0F);
		bw.Write((uint)0x02);
		bw.Write((ushort)Modules.Count);
		bw.Write((ushort)0x13);
		bw.Write((uint)0x02);
		bw.Write((ushort)0xFFFF);

		foreach (var module in Modules)
		{
			WriteModuleRecord(bw, module);
		}

		bw.Write((ushort)0x10);             //Terminator
		bw.Write((uint)0);

		return VBACompression.CompressPart(((MemoryStream)bw.BaseStream).ToArray());
	}

	private void WriteProjectInformation(BinaryWriter bw)
	{
		/****** PROJECTINFORMATION Record ******/
		bw.Write((ushort)1);        //ID
		bw.Write((uint)4);          //Size
		bw.Write((uint)SystemKind); //SysKind

		bw.Write((ushort)2);        //ID
		bw.Write((uint)4);          //Size
		bw.Write((uint)Lcid);       //Lcid

		bw.Write((ushort)0x14);     //ID
		bw.Write((uint)4);          //Size
		bw.Write((uint)LcidInvoke); //Lcid Invoke

		bw.Write((ushort)3);        //ID
		bw.Write((uint)2);          //Size
		bw.Write((ushort)CodePage);   //Codepage

		//ProjectName
		bw.Write((ushort)4);                                            //ID
		bw.Write((uint)Name.Length);                             //Size
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Name)); //Project Name

		//Description
		bw.Write((ushort)5);                                            //ID
		bw.Write((uint)Description.Length);                             //Size
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Description)); //Project Name
		bw.Write((ushort)0x40);                                           //ID
		bw.Write((uint)Description.Length * 2);                           //Size
		bw.Write(Encoding.Unicode.GetBytes(Description));               //Project Description

		//Helpfiles
		bw.Write((ushort)6);                                           //ID
		bw.Write((uint)HelpFile1.Length);                              //Size
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(HelpFile1));  //HelpFile1            
		bw.Write((ushort)0x3D);                                           //ID
		bw.Write((uint)HelpFile2.Length);                              //Size
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(HelpFile2));  //HelpFile2

		//Help context id
		bw.Write((ushort)7);            //ID
		bw.Write((uint)4);              //Size
		bw.Write((uint)HelpContextID);  //Help context id

		//Libflags
		bw.Write((ushort)8);            //ID
		bw.Write((uint)4);              //Size
		bw.Write((uint)0);  //Help context id

		//Vba Version
		bw.Write((ushort)9);            //ID
		bw.Write((uint)4);              //Reserved
		bw.Write((uint)MajorVersion);   //Reserved
		bw.Write((ushort)MinorVersion); //Help context id

		//Constants
		bw.Write((ushort)0x0C);           //ID
		bw.Write((uint)Constants.Length);              //Size
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Constants));              //Help context id
		bw.Write((ushort)0x3C);                                           //ID
		bw.Write((uint)Constants.Length / 2);                              //Size
		bw.Write(Encoding.Unicode.GetBytes(Constants));  //HelpFile2
	}

	private void WriteModuleRecord(BinaryWriter bw, ExcelVBAModule module)
	{
		bw.Write((ushort)0x19);
		bw.Write((uint)module.Name.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));     //Name

		bw.Write((ushort)0x47);
		bw.Write((uint)module.Name.Length * 2);
		bw.Write(Encoding.Unicode.GetBytes(module.Name));                   //Name

		bw.Write((ushort)0x1A);
		bw.Write((uint)module.Name.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));     //Stream Name  

		bw.Write((ushort)0x32);
		bw.Write((uint)module.Name.Length * 2);
		bw.Write(Encoding.Unicode.GetBytes(module.Name));                   //Stream Name

		module.Description = module.Description ?? "";
		bw.Write((ushort)0x1C);
		bw.Write((uint)module.Description.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Description));     //Description

		bw.Write((ushort)0x48);
		bw.Write((uint)module.Description.Length * 2);
		bw.Write(Encoding.Unicode.GetBytes(module.Description));                   //Description

		bw.Write((ushort)0x31);
		bw.Write((uint)4);
		bw.Write((uint)0);                              //Module Stream Offset (No PerformanceCache)

		bw.Write((ushort)0x1E);
		bw.Write((uint)4);
		bw.Write((uint)module.HelpContext);            //Help context ID

		bw.Write((ushort)0x2C);
		bw.Write((uint)2);
		bw.Write((ushort)0xFFFF);            //Help context ID

		bw.Write((ushort)(module.Type == eModuleType.Module ? 0x21 : 0x22));
		bw.Write((uint)0);

		if (module.ReadOnly)
		{
			bw.Write((ushort)0x25);
			bw.Write((uint)0);              //Readonly
		}

		if (module.Private)
		{
			bw.Write((ushort)0x28);
			bw.Write((uint)0);              //Private
		}

		bw.Write((ushort)0x2B);             //Terminator
		bw.Write((uint)0);
	}

	private void WriteNameReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		//Name record
		bw.Write((ushort)0x16);                                             //ID
		bw.Write((uint)reference.Name.Length);                              //Size
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Name));  //HelpFile1
		bw.Write((ushort)0x3E);                                             //ID
		bw.Write((uint)reference.Name.Length * 2);                            //Size
		bw.Write(Encoding.Unicode.GetBytes(reference.Name));                //HelpFile2
	}
	private void WriteControlReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		WriteOrginalReference(bw, reference);

		bw.Write((ushort)0x2F);
		var controlRef = (ExcelVbaReferenceControl)reference;
		bw.Write((uint)(4 + controlRef.LibIdTwiddled.Length + 4 + 2));    // Size of SizeOfLibidTwiddled, LibidTwiddled, Reserved1, and Reserved2.
		bw.Write((uint)controlRef.LibIdTwiddled.Length);                              //Size            
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(controlRef.LibIdTwiddled));  //LibID
		bw.Write((uint)0);      //Reserved1
		bw.Write((ushort)0);    //Reserved2
		WriteNameReference(bw, reference);  //Name record again
		bw.Write((ushort)0x30); //Reserved3
		bw.Write((uint)(4 + controlRef.LibIdExternal.Length + 4 + 2 + 16 + 4));    //Size of SizeOfLibidExtended, LibidExtended, Reserved4, Reserved5, OriginalTypeLib, and Cookie
		bw.Write((uint)controlRef.LibIdExternal.Length);                              //Size            
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(controlRef.LibIdExternal));  //LibID
		bw.Write((uint)0);      //Reserved4
		bw.Write((ushort)0);    //Reserved5
		bw.Write(controlRef.OriginalTypeLib.ToByteArray());
		bw.Write((uint)controlRef.Cookie);      //Cookie
	}

	private void WriteOrginalReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		bw.Write((ushort)0x33);
		bw.Write((uint)reference.Libid.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Libid));  //LibID
	}
	private void WriteProjectReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		bw.Write((ushort)0x0E);
		var projRef = (ExcelVbaReferenceProject)reference;
		bw.Write((uint)(4 + projRef.Libid.Length + 4 + projRef.LibIdRelative.Length + 4 + 2));
		bw.Write((uint)projRef.Libid.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(projRef.Libid));  //LibAbsolute
		bw.Write((uint)projRef.LibIdRelative.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(projRef.LibIdRelative));  //LibIdRelative
		bw.Write(projRef.MajorVersion);
		bw.Write(projRef.MinorVersion);
	}

	private void WriteRegisteredReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		bw.Write((ushort)0x0D);
		bw.Write((uint)(4 + reference.Libid.Length + 4 + 2));
		bw.Write((uint)reference.Libid.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Libid));  //LibID            
		bw.Write((uint)0);      //Reserved1
		bw.Write((ushort)0);    //Reserved2
	}

	private byte[] CreateProjectwmStream()
	{
		BinaryWriter bw = new(new MemoryStream());

		foreach (var module in Modules)
		{
			bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));     //Name
			bw.Write((byte)0); //Null
			bw.Write(Encoding.Unicode.GetBytes(module.Name));                   //Name
			bw.Write((ushort)0); //Null
		}

		bw.Write((ushort)0); //Null
		return ((MemoryStream)bw.BaseStream).ToArray();
	}
	private byte[] CreateProjectStream()
	{
		StringBuilder sb = new();
		sb.AppendFormat("ID=\"{0}\"\r\n", ProjectID);
		foreach (var module in Modules)
		{
			if (module.Type == eModuleType.Document)
			{
				sb.AppendFormat("Document={0}/&H00000000\r\n", module.Name);
			}
			else if (module.Type == eModuleType.Module)
			{
				sb.AppendFormat("Module={0}\r\n", module.Name);
			}
			else if (module.Type == eModuleType.Class)
			{
				sb.AppendFormat("Class={0}\r\n", module.Name);
			}
			else
			{
				//Designer
				sb.AppendFormat("Package={0}\r\n", module.ClassID);
				sb.AppendFormat("BaseClass={0}\r\n", module.Name);
			}
		}

		if (HelpFile1 != "")
		{
			sb.AppendFormat("HelpFile={0}\r\n", HelpFile1);
		}

		sb.AppendFormat("Name=\"{0}\"\r\n", Name);
		sb.AppendFormat("HelpContextID={0}\r\n", HelpContextID);

		if (!string.IsNullOrEmpty(Description))
		{
			sb.AppendFormat("Description=\"{0}\"\r\n", Description);
		}

		sb.Append("VersionCompatible32=\"393222000\"\r\n");

		sb.AppendFormat("CMG=\"{0}\"\r\n", WriteProtectionStat());
		sb.AppendFormat("DPB=\"{0}\"\r\n", WritePassword());
		sb.AppendFormat("GC=\"{0}\"\r\n\r\n", WriteVisibilityState());

		sb.Append("[Host Extender Info]\r\n");
		sb.Append("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");
		sb.Append("\r\n");
		sb.Append("[Workspace]\r\n");
		foreach (var module in Modules)
		{
			sb.AppendFormat("{0}=0, 0, 0, 0, C \r\n", module.Name);
		}

		var s = sb.ToString();
		return Encoding.GetEncoding(CodePage).GetBytes(s);
	}
	private string WriteProtectionStat()
	{
		var stat = (_protection.UserProtected ? 1 : 0) |
				 (_protection.HostProtected ? 2 : 0) |
				 (_protection.VbeProtected ? 4 : 0);

		return Encrypt(BitConverter.GetBytes(stat));
	}
	private string WritePassword()
	{
		var nullBits = new byte[3];
		var nullKey = new byte[4];
		var nullHash = new byte[20];
		if (Protection.PasswordKey == null)
		{
			return Encrypt([0]);
		}
		else
		{
			Array.Copy(Protection.PasswordKey, nullKey, 4);
			Array.Copy(Protection.PasswordHash, nullHash, 20);

			//Set Null bits
			for (var i = 0; i < 24; i++)
			{
				var bit = (byte)(128 >> (i % 8));
				if (i < 4)
				{
					if (nullKey[i] == 0)
					{
						nullKey[i] = 1;
					}
					else
					{
						nullBits[0] |= bit;
					}
				}
				else
				{
					if (nullHash[i - 4] == 0)
					{
						nullHash[i - 4] = 1;
					}
					else
					{
						var byteIndex = (i - i % 8) / 8;
						nullBits[byteIndex] |= bit;
					}
				}
			}
			//Write the Password Hash Data Structure (2.4.4.1)
			BinaryWriter bw = new(new MemoryStream());
			bw.Write((byte)0xFF);
			bw.Write(nullBits);
			bw.Write(nullKey);
			bw.Write(nullHash);
			bw.Write((byte)0);
			return Encrypt(((MemoryStream)bw.BaseStream).ToArray());
		}
	}
	private string WriteVisibilityState() => Encrypt([(byte)(Protection.VisibilityState ? 0xFF : 0)]);
	#endregion
}
