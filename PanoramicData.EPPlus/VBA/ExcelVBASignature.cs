/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
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
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace OfficeOpenXml.VBA;

/// <summary>
/// The code signature properties of the project
/// </summary>
public class ExcelVbaSignature
{
	const string schemaRelVbaSignature = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
	readonly Packaging.ZipPackagePart _vbaPart;
	internal ExcelVbaSignature(Packaging.ZipPackagePart vbaPart)
	{
		_vbaPart = vbaPart;
		GetSignature();
	}
	private void GetSignature()
	{
		if (_vbaPart == null)
		{
			return;
		}

		var rel = _vbaPart.GetRelationshipsByType(schemaRelVbaSignature).FirstOrDefault();
		if (rel != null)
		{
			Uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
			Part = _vbaPart.Package.GetPart(Uri);

			var stream = Part.GetStream();
			BinaryReader br = new(stream);
			var cbSignature = br.ReadUInt32();
			// signatureOffset, cbSigningCertStore, certStoreOffset, cbProjectName, projectNameOffset, fTimestamp, cbTimestampUrl, timestampUrlOffset
			SkipUInt32(br, 8);
			var signature = br.ReadBytes((int)cbSignature);
			// version, fileType
			SkipUInt32(br, 2);
			ReadCertificateEntries(br);
			// endel1, endel2, rgchProjectNameBuffer, rgchTimestampBuffer
			SkipUInt32(br, 2);
			_ = br.ReadUInt16();
			_ = br.ReadUInt16();
			Verifier = new EnvelopedCms();
			Verifier.Decode(signature);
		}
		else
		{
			Certificate = null;
			Verifier = null;
		}
	}
	private static void SkipUInt32(BinaryReader br, int count)
	{
		for (var i = 0; i < count; i++)
		{
			_ = br.ReadUInt32();
		}
	}

	private void ReadCertificateEntries(BinaryReader br)
	{
		var id = br.ReadUInt32();
		while (id != 0)
		{
			// encodingType
			_ = br.ReadUInt32();
			var length = br.ReadUInt32();
			if (length > 0)
			{
				var value = br.ReadBytes((int)length);
				if (id == 0x20)
				{
					Certificate = X509CertificateLoader.LoadCertificate(value);
				}
			}

			id = br.ReadUInt32();
		}
	}
	internal void Save(ExcelVbaProject proj)
	{
		if (Certificate == null)
		{
			return;
		}

		//No signature. Remove any Signature part
		if (Certificate.HasPrivateKey == false && !TryUseStoreCertificate())
		{
			RemoveSignaturePart();
			return;
		}

		var certStore = GetCertStore();
		var cert = SignProject(proj);
		var b = BuildSignatureStream(cert, certStore);
		EnsureSignaturePart(proj);
		Part.GetStream(FileMode.Create).Write(b, 0, b.Length);
	}

	private bool TryUseStoreCertificate()
	{
		var storeCert = GetCertFromStore(StoreLocation.CurrentUser);
		storeCert ??= GetCertFromStore(StoreLocation.LocalMachine);

		if (storeCert != null && storeCert.HasPrivateKey == true)
		{
			Certificate = storeCert;
			return true;
		}

		return false;
	}

	private void RemoveSignaturePart()
	{
		foreach (var r in Part.GetRelationships())
		{
			Part.DeleteRelationship(r.Id);
		}

		Part.Package.DeletePart(Part.Uri);
	}

	private static byte[] BuildSignatureStream(byte[] cert, byte[] certStore)
	{
		var ms = new MemoryStream();
		var bw = new BinaryWriter(ms);
		bw.Write((uint)cert.Length);
		bw.Write((uint)44);                  //?? 36 ref inside cert ??
		bw.Write((uint)certStore.Length);    //cbSigningCertStore
		bw.Write((uint)(cert.Length + 44));  //certStoreOffset
		bw.Write((uint)0);                   //cbProjectName
		bw.Write((uint)(cert.Length + certStore.Length + 44));    //projectNameOffset
		bw.Write((uint)0);    //fTimestamp
		bw.Write((uint)0);    //cbTimestampUrl
		bw.Write((uint)(cert.Length + certStore.Length + 44 + 2));    //timestampUrlOffset
		bw.Write(cert);
		bw.Write(certStore);
		bw.Write((ushort)0);//rgchProjectNameBuffer
		bw.Write((ushort)0);//rgchTimestampBuffer
		bw.Write((ushort)0);
		bw.Flush();
		return ms.ToArray();
	}

	private void EnsureSignaturePart(ExcelVbaProject proj)
	{
		var rel = proj.Part.GetRelationshipsByType(schemaRelVbaSignature).FirstOrDefault();
		if (Part == null)
		{
			if (rel != null)
			{
				Uri = rel.TargetUri;
				Part = proj._pck.GetPart(rel.TargetUri);
			}
			else
			{
				Uri = new Uri("/xl/vbaProjectSignature.bin", UriKind.Relative);
				Part = proj._pck.CreatePart(Uri, ExcelPackage.schemaVBASignature);
			}
		}

		if (rel == null)
		{
			proj.Part.CreateRelationship(UriHelper.ResolvePartUri(proj.Uri, Uri), Packaging.TargetMode.Internal, schemaRelVbaSignature);
		}
	}

	private X509Certificate2 GetCertFromStore(StoreLocation loc)
	{
		try
		{
			X509Store store = new(StoreName.My, loc);
			store.Open(OpenFlags.ReadOnly);
			try
			{
				var storeCert = store.Certificates.Find(
								X509FindType.FindByThumbprint,
								Certificate.Thumbprint,
								true
								).OfType<X509Certificate2>().FirstOrDefault();
				return storeCert;
			}
			finally
			{
#if Core
                        store.Dispose();
#endif
				store.Close();
			}
		}
		catch
		{
			return null;
		}
	}

	private byte[] GetCertStore()
	{
		var ms = new MemoryStream();
		var bw = new BinaryWriter(ms);

		bw.Write((uint)0); //Version
		bw.Write((uint)0x54524543); //fileType

		//SerializedCertificateEntry
		var certData = Certificate.RawData;
		bw.Write((uint)0x20);
		bw.Write((uint)1);
		bw.Write((uint)certData.Length);
		bw.Write(certData);

		//EndElementMarkerEntry
		bw.Write((uint)0);
		bw.Write((ulong)0);

		bw.Flush();
		return ms.ToArray();
	}

	private static void WriteProp(BinaryWriter bw, int id, byte[] data)
	{
		bw.Write((uint)id);
		bw.Write((uint)1);
		bw.Write((uint)data.Length);
		bw.Write(data);
	}
	internal byte[] SignProject(ExcelVbaProject proj)
	{
		if (!Certificate.HasPrivateKey)
		{
			Certificate = null;
			return null;
		}

		var hash = GetContentHash(proj);

		BinaryWriter bw = new(new MemoryStream());
		bw.Write((byte)0x30); //Constructed Type 
		bw.Write((byte)0x32); //Total length
		bw.Write((byte)0x30); //Constructed Type 
		bw.Write((byte)0x0E); //Length SpcIndirectDataContent
		bw.Write((byte)0x06); //Oid Tag Indentifier 
		bw.Write((byte)0x0A); //Lenght OId
		bw.Write([0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1D]); //Encoded Oid 1.3.6.1.4.1.311.2.1.29
		bw.Write((byte)0x04);   //Octet String Tag Identifier
		bw.Write((byte)0x00);   //Zero length

		bw.Write((byte)0x30); //Constructed Type (DigestInfo)
		bw.Write((byte)0x20); //Length DigestInfo
		bw.Write((byte)0x30); //Constructed Type (Algorithm)
		bw.Write((byte)0x0C); //length AlgorithmIdentifier
		bw.Write((byte)0x06); //Oid Tag Indentifier 
		bw.Write((byte)0x08); //Lenght OId
		bw.Write([0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05]); //Encoded Oid for 1.2.840.113549.2.5 (AlgorithmIdentifier MD5)
		bw.Write((byte)0x05);   //Null type identifier
		bw.Write((byte)0x00);   //Null length
		bw.Write((byte)0x04);   //Octet String Identifier
		bw.Write((byte)hash.Length);   //Hash length
		bw.Write(hash);                //Content hash

		var contentInfo = new ContentInfo(((MemoryStream)bw.BaseStream).ToArray());
		contentInfo.ContentType.Value = "1.3.6.1.4.1.311.2.1.4";
		Verifier = new EnvelopedCms(contentInfo);
		var r = new CmsRecipient(Certificate);
		Verifier.Encrypt(r);
		return Verifier.Encode();
	}

	private static byte[] GetContentHash(ExcelVbaProject proj)
	{
		//MS-OVBA 2.4.2
		var enc = System.Text.Encoding.GetEncoding(proj.CodePage);
		BinaryWriter bw = new(new MemoryStream());
		bw.Write(enc.GetBytes(proj.Name));
		bw.Write(enc.GetBytes(proj.Constants));
		foreach (var reference in proj.References)
		{
			if (reference.ReferenceRecordID == 0x0D)
			{
				bw.Write((byte)0x7B);
			}

			if (reference.ReferenceRecordID == 0x0E)
			{
				foreach (var b in BitConverter.GetBytes((uint)reference.Libid.Length))  //Length will never be an UInt with 4 bytes that aren't 0 (> 0x00FFFFFF), so no need for the rest of the properties.
				{
					if (b != 0)
					{
						bw.Write(b);
					}
					else
					{
						break;
					}
				}
			}
		}

		foreach (var module in proj.Modules)
		{
			var lines = module.Code.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
			foreach (var line in lines)
			{
				if (!line.StartsWith("attribute", StringComparison.OrdinalIgnoreCase))
				{
					bw.Write(enc.GetBytes(line));
				}
			}
		}

		var buffer = (bw.BaseStream as MemoryStream).ToArray();
		return System.Security.Cryptography.MD5.HashData(buffer);
	}
	/// <summary>
	/// The certificate to sign the VBA project.
	/// <remarks>
	/// This certificate must have a private key.
	/// There is no validation that the certificate is valid for codesigning, so make sure it's valid to sign Excel files (Excel 2010 is more strict that prior versions).
	/// </remarks>
	/// </summary>
	public X509Certificate2 Certificate { get; set; }
	/// <summary>
	/// The verifier
	/// </summary>
	public EnvelopedCms Verifier { get; internal set; }
	internal CompoundDocument Signature { get; set; }
	internal Packaging.ZipPackagePart Part { get; set; }
	internal Uri Uri { get; private set; }
}
