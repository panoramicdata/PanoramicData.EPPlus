using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.VBA;

public partial class ExcelVbaProject
{
	#region Read Project
	private void GetProject()
	{

		var stream = Part.GetStream();
		byte[] vba;
		vba = new byte[stream.Length];
		stream.Read(vba, 0, (int)stream.Length);
		Document = new CompoundDocument(vba);

		ReadDirStream();
		ProjectStreamText = Encoding.GetEncoding(CodePage).GetString(Document.Storage.DataStreams["PROJECT"]);
		ReadModules();
		ReadProjectProperties();
	}
	private void ReadModules()
	{
		foreach (var modul in Modules)
		{
			var stream = Document.Storage.SubStorage["VBA"].DataStreams[modul.streamName];
			var byCode = VBACompression.DecompressPart(stream, (int)modul.ModuleOffset);
			var code = Encoding.GetEncoding(CodePage).GetString(byCode);
			var pos = 0;
			while (pos + 9 < code.Length && code.Substring(pos, 9) == "Attribute")
			{
				var linePos = code.IndexOf("\r\n", pos, StringComparison.CurrentCulture);
				string[] lineSplit;
				if (linePos > 0)
				{
					lineSplit = code.Substring(pos + 9, linePos - pos - 9).Split('=');
				}
				else
				{
					lineSplit = code[(pos + 9)..].Split(['='], 1);
				}

				if (lineSplit.Length > 1)
				{
					lineSplit[1] = lineSplit[1].Trim();
					var attr =
						new ExcelVbaModuleAttribute
						{
							Name = lineSplit[0].Trim(),
							DataType = lineSplit[1].StartsWith("\"") ? eAttributeDataType.String : eAttributeDataType.NonString,
							Value = lineSplit[1].StartsWith("\"") ? lineSplit[1][1..^1] : lineSplit[1]
						};
					modul.Attributes._list.Add(attr);
				}

				pos = linePos + 2;
			}

			modul.Code = code[pos..];
		}
	}

	private void ReadProjectProperties()
	{
		_protection = new ExcelVbaProtection(this);
		var prevPackage = "";
		var lines = Regex.Split(ProjectStreamText, "\r\n");
		foreach (var line in lines)
		{
			if (!line.StartsWith("["))
			{
				var split = line.Split('=');
				if (split.Length > 1 && split[1].Length > 1 && split[1].StartsWith("\"")) //Remove any double qouates
				{
					split[1] = split[1][1..^1];
				}

				if (!ReadModuleProperty(split, ref prevPackage))
				{
					ReadProtectionProperty(split);
				}
			}
		}
	}

	private bool ReadModuleProperty(string[] split, ref string prevPackage)
	{
		switch (split[0])
		{
			case "Document":
				var mn = split[1][..split[1].IndexOf("/&H")];
				Modules[mn].Type = eModuleType.Document;
				return true;
			case "Package":
				prevPackage = split[1];
				return true;
			case "BaseClass":
				Modules[split[1]].Type = eModuleType.Designer;
				Modules[split[1]].ClassID = prevPackage;
				return true;
			case "Module":
				Modules[split[1]].Type = eModuleType.Module;
				return true;
			case "Class":
				Modules[split[1]].Type = eModuleType.Class;
				return true;
			default:
				return false;
		}
	}

	private void ReadProtectionProperty(string[] split)
	{
		switch (split[0])
		{
			case "ID":
				ProjectID = split[1];
				break;
			case "CMG":
				var cmg = Decrypt(split[1]);
				_protection.UserProtected = (cmg[0] & 1) != 0;
				_protection.HostProtected = (cmg[0] & 2) != 0;
				_protection.VbeProtected = (cmg[0] & 4) != 0;
				break;
			case "DPB":
				ReadPasswordHash(Decrypt(split[1]));
				break;
			case "GC":
				_protection.VisibilityState = Decrypt(split[1])[0] == 0xFF;
				break;
			default:
				// HelpFile, Name, HelpContextID, Description and VersionCompatible32 are not read back
				break;
		}
	}

	private void ReadPasswordHash(byte[] dpb)
	{
		if (dpb.Length < 28)
		{
			return;
		}

		var flags = new byte[3];
		Array.Copy(dpb, 1, flags, 0, 3);
		var keyNoNulls = new byte[4];
		_protection.PasswordKey = new byte[4];
		Array.Copy(dpb, 4, keyNoNulls, 0, 4);
		var hashNoNulls = new byte[20];
		_protection.PasswordHash = new byte[20];
		Array.Copy(dpb, 8, hashNoNulls, 0, 20);
		//Handle 0x00 bitwise 2.4.4.3
		for (var i = 0; i < 24; i++)
		{
			var bit = 128 >> (i % 8);
			if (i < 4)
			{
				_protection.PasswordKey[i] = (flags[0] & bit) == 0 ? (byte)0 : keyNoNulls[i];
			}
			else
			{
				var flagIndex = (i - i % 8) / 8;
				_protection.PasswordHash[i - 4] = (flags[flagIndex] & bit) == 0 ? (byte)0 : hashNoNulls[i - 4];
			}
		}
	}

	/// <summary>
	/// 2.4.3.3 Decryption
	/// </summary>
	/// <param name="value">Byte hex string</param>
	/// <returns>The decrypted value</returns>
	private static byte[] Decrypt(string value)
	{
		var enc = GetByte(value);
		var dec = new byte[(value.Length - 1)];
		byte seed, ignoredLength;
		seed = enc[0];
		dec[0] = (byte)(enc[1] ^ seed);
		dec[1] = (byte)(enc[2] ^ seed);
		for (var i = 2; i < enc.Length - 1; i++)
		{
			dec[i] = (byte)(enc[i + 1] ^ (enc[i - 1] + dec[i - 1]));
		}

		ignoredLength = (byte)((seed & 6) / 2);
		var datalength = BitConverter.ToInt32(dec, ignoredLength + 2);
		var data = new byte[datalength];
		Array.Copy(dec, 6 + ignoredLength, data, 0, datalength);
		return data;
	}
	/// <summary>
	/// 2.4.3.2 Encryption
	/// </summary>
	/// <param name="value"></param>
	/// <returns>Byte hex string</returns>
	private string Encrypt(byte[] value)
	{
		var seed = new byte[1];
		var rn = RandomNumberGenerator.Create();
		rn.GetBytes(seed);
		BinaryWriter br = new(new MemoryStream());
		var enc = new byte[value.Length + 10];
		enc[0] = seed[0];
		enc[1] = (byte)(2 ^ seed[0]);

		byte projKey = 0;

		foreach (var c in ProjectID)
		{
			projKey += (byte)c;
		}

		enc[2] = (byte)(projKey ^ seed[0]);
		var ignoredLength = (seed[0] & 6) / 2;
		for (var i = 0; i < ignoredLength; i++)
		{
			br.Write(seed[0]);
		}

		br.Write(value.Length);
		br.Write(value);

		var pos = 3;
		var pb = projKey;
		foreach (var b in ((MemoryStream)br.BaseStream).ToArray())
		{
			enc[pos] = (byte)(b ^ (enc[pos - 2] + pb));
			pos++;
			pb = b;
		}

		return GetString(enc, pos - 1);
	}
	private static string GetString(byte[] value, int max)
	{
		var ret = new StringBuilder();
		for (var i = 0; i <= max; i++)
		{
			ret.Append(value[i].ToString("x2"));
		}

		return ret.ToString().ToUpperInvariant();
	}
	private static byte[] GetByte(string value)
	{
		var ret = new byte[value.Length / 2];
		for (var i = 0; i < ret.Length; i++)
		{
			ret[i] = byte.Parse(value.Substring(i * 2, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
		}

		return ret;
	}
	private void ReadDirStream()
	{
		var dir = VBACompression.DecompressPart(Document.Storage.SubStorage["VBA"].DataStreams["dir"]);
		MemoryStream ms = new(dir);
		BinaryReader br = new(ms);
		var state = new DirStreamState();
		while (br.BaseStream.Position < br.BaseStream.Length)
		{
			var id = br.ReadUInt16();
			var size = br.ReadUInt32();
			//Terminator
			if (id == 0x10)
			{
				break;
			}

			if (!ReadProjectInfoRecord(br, id, size)
				&& !ReadProjectVersionRecord(br, id, size)
				&& !ReadReferenceRecord(br, id, size, state)
				&& !ReadModuleRecord(br, id, size, state))
			{
				ReadModuleFlagRecord(br, id, state);
			}
		}
	}

	private sealed class DirStreamState
	{
		internal ExcelVbaReference CurrentRef { get; set; }
		internal string ReferenceName { get; set; } = "";
		internal ExcelVBAModule CurrentModule { get; set; }
	}

	private bool ReadProjectInfoRecord(BinaryReader br, ushort id, uint size)
	{
		switch (id)
		{
			case 0x01:
				SystemKind = (eSyskind)br.ReadUInt32();
				return true;
			case 0x02:
				Lcid = (int)br.ReadUInt32();
				return true;
			case 0x03:
				CodePage = (int)br.ReadUInt16();
				return true;
			case 0x04:
				Name = GetString(br, size);
				return true;
			case 0x05:
				Description = GetUnicodeString(br, size);
				return true;
			case 0x06:
				HelpFile1 = GetString(br, size);
				return true;
			case 0x3D:
				HelpFile2 = GetString(br, size);
				return true;
			default:
				return false;
		}
	}

	private bool ReadProjectVersionRecord(BinaryReader br, ushort id, uint size)
	{
		switch (id)
		{
			case 0x07:
				HelpContextID = (int)br.ReadUInt32();
				return true;
			case 0x08:
				LibFlags = (int)br.ReadUInt32();
				return true;
			case 0x09:
				MajorVersion = (int)br.ReadUInt32();
				MinorVersion = (int)br.ReadUInt16();
				return true;
			case 0x0C:
				Constants = GetUnicodeString(br, size);
				return true;
			case 0x14:
				LcidInvoke = (int)br.ReadUInt32();
				return true;
			case 0x0F:
			case 0x13:
				// Module count and cookie are not kept
				_ = br.ReadUInt16();
				return true;
			default:
				return false;
		}
	}

	private bool ReadReferenceRecord(BinaryReader br, ushort id, uint size, DirStreamState state)
	{
		switch (id)
		{
			case 0x16:
				state.ReferenceName = GetUnicodeString(br, size);
				return true;
			case 0x0D:
				ReadRegisteredReference(br, id, state.ReferenceName);
				return true;
			case 0x0E:
				ReadProjectReference(br, id, state.ReferenceName);
				return true;
			case 0x33:
				state.CurrentRef = new ExcelVbaReferenceControl
				{
					ReferenceRecordID = id,
					Name = state.ReferenceName,
					Libid = GetString(br, size)
				};
				References.Add(state.CurrentRef);
				return true;
			case 0x30:
				ReadExtendedReference(br, (ExcelVbaReferenceControl)state.CurrentRef);
				return true;
			case 0x2F:
				ReadTwiddledReference(br, id, (ExcelVbaReferenceControl)state.CurrentRef);
				return true;
			default:
				return false;
		}
	}

	private void ReadRegisteredReference(BinaryReader br, ushort id, string referenceName)
	{
		var sizeLibID = br.ReadUInt32();
		var regRef = new ExcelVbaReference
		{
			Name = referenceName,
			ReferenceRecordID = id,
			Libid = GetString(br, sizeLibID)
		};
		// Reserved1, Reserved2
		_ = br.ReadUInt32();
		_ = br.ReadUInt16();
		References.Add(regRef);
	}

	private void ReadProjectReference(BinaryReader br, ushort id, string referenceName)
	{
		var projRef = new ExcelVbaReferenceProject
		{
			ReferenceRecordID = id,
			Name = referenceName
		};
		var sizeLibID = br.ReadUInt32();
		projRef.Libid = GetString(br, sizeLibID);
		sizeLibID = br.ReadUInt32();
		projRef.LibIdRelative = GetString(br, sizeLibID);
		projRef.MajorVersion = br.ReadUInt32();
		projRef.MinorVersion = br.ReadUInt16();
		References.Add(projRef);
	}

	private void ReadExtendedReference(BinaryReader br, ExcelVbaReferenceControl extRef)
	{
		var sizeExt = br.ReadUInt32();
		extRef.LibIdExternal = GetString(br, sizeExt);
		// Reserved4, Reserved5
		_ = br.ReadUInt32();
		_ = br.ReadUInt16();
		extRef.OriginalTypeLib = new Guid(br.ReadBytes(16));
		extRef.Cookie = br.ReadUInt32();
	}

	private void ReadTwiddledReference(BinaryReader br, ushort id, ExcelVbaReferenceControl contrRef)
	{
		contrRef.ReferenceRecordID = id;
		var sizeTwiddled = br.ReadUInt32();
		contrRef.LibIdTwiddled = GetString(br, sizeTwiddled);
		// Reserved1, Reserved2
		_ = br.ReadUInt32();
		_ = br.ReadUInt16();
	}

	private bool ReadModuleRecord(BinaryReader br, ushort id, uint size, DirStreamState state)
	{
		switch (id)
		{
			case 0x19:
				state.CurrentModule = new ExcelVBAModule
				{
					Name = GetUnicodeString(br, size)
				};
				Modules.Add(state.CurrentModule);
				return true;
			case 0x1A:
				state.CurrentModule.streamName = GetUnicodeString(br, size);
				return true;
			case 0x1C:
				state.CurrentModule.Description = GetUnicodeString(br, size);
				return true;
			case 0x1E:
				state.CurrentModule.HelpContext = (int)br.ReadUInt32();
				return true;
			default:
				return false;
		}
	}

	private static void ReadModuleFlagRecord(BinaryReader br, ushort id, DirStreamState state)
	{
		switch (id)
		{
			case 0x2C:
				state.CurrentModule.Cookie = br.ReadUInt16();
				break;
			case 0x31:
				state.CurrentModule.ModuleOffset = br.ReadUInt32();
				break;
			case 0x25:
				state.CurrentModule.ReadOnly = true;
				break;
			case 0x28:
				state.CurrentModule.Private = true;
				break;
			default:
				// Module type (0x21, 0x22), terminator (0x2B) and unknown records carry no data to keep
				break;
		}
	}
	#endregion
}
