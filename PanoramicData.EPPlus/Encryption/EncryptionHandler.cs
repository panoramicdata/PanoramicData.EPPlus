﻿/*******************************************************************************
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
 * ******************************************************************************
 * Jan Källman		    Added       		        2013-01-05
 *******************************************************************************/
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Text;

namespace OfficeOpenXml.Encryption;

/// <summary>
/// Handles encrypted Excel documents 
/// </summary>
internal class EncryptedPackageHandler
{
	/// <summary>
	/// Read the package from the OLE document and decrypt it using the supplied password
	/// </summary>
	/// <param name="fi">The file</param>
	/// <param name="encryption"></param>
	/// <returns></returns>
	internal MemoryStream DecryptPackage(FileInfo fi, ExcelEncryption encryption)
	{
		if (CompoundDocument.IsCompoundDocument(fi))
		{
			CompoundDocument doc = new(fi);

			MemoryStream ret = null;
			ret = GetStreamFromPackage(doc, encryption);
			return ret;
		}
		else
		{
			throw (new InvalidDataException(string.Format("File {0} is not an encrypted package", fi.FullName)));
		}
	}
	//Helpmethod to output the streams in the storage
	//private void WriteDoc(CompoundDocument.StoragePart storagePart, string p)
	//{
	//    foreach (var store in storagePart.SubStorage)
	//    {
	//        string sdir=p + store.Key.Replace((char)6,'x') + "\\";
	//        Directory.CreateDirectory(sdir);
	//        WriteDoc(store.Value, sdir);
	//    }
	//    foreach (var str in storagePart.DataStreams)
	//    {
	//        File.WriteAllBytes(p + str.Key.Replace((char)6, 'x') + ".bin", str.Value);
	//    }
	//}
	/// <summary>
	/// Read the package from the OLE document and decrypt it using the supplied password
	/// </summary>
	/// <param name="stream">The memory stream. </param>
	/// <param name="encryption">The encryption object from the Package</param>
	/// <returns></returns>
	internal MemoryStream DecryptPackage(MemoryStream stream, ExcelEncryption encryption)
	{
		try
		{
			if (CompoundDocument.IsCompoundDocument(stream))
			{
				var doc = new CompoundDocument(stream);
				return GetStreamFromPackage(doc, encryption);
			}
			else
			{
				throw (new InvalidDataException("The stream is not an valid/supported encrypted document."));
			}
		}
		catch// (Exception ex)
		{
			throw;
		}

	}
	/// <summary>
	/// Encrypts a package
	/// </summary>
	/// <param name="package">The package as a byte array</param>
	/// <param name="encryption">The encryption info from the workbook</param>
	/// <returns></returns>
	internal MemoryStream EncryptPackage(byte[] package, ExcelEncryption encryption)
	{
		if (encryption.Version == EncryptionVersion.Standard) //Standard encryption
		{
			return EncryptPackageBinary(package, encryption);
		}
		else if (encryption.Version == EncryptionVersion.Agile) //Agile encryption
		{
			return EncryptPackageAgile(package, encryption);
		}

		throw (new ArgumentException("Unsupported encryption version."));
	}
	private MemoryStream EncryptPackageAgile(byte[] package, ExcelEncryption encryption)
	{
		var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
		xml += "<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\" xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">";
		xml += "<keyData saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"\"/>";
		xml += "<dataIntegrity encryptedHmacKey=\"\" encryptedHmacValue=\"\"/>";
		xml += "<keyEncryptors>";
		xml += "<keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">";
		xml += "<p:encryptedKey spinCount=\"100000\" saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"\" encryptedVerifierHashInput=\"\" encryptedVerifierHashValue=\"\" encryptedKeyValue=\"\" />";
		xml += "</keyEncryptor></keyEncryptors></encryption>";

		var encryptionInfo = new EncryptionInfoAgile();
		encryptionInfo.ReadFromXml(xml);
		var encr = encryptionInfo.KeyEncryptors[0];
		var rnd = RandomNumberGenerator.Create();

		var s = new byte[16];
		rnd.GetBytes(s);
		encryptionInfo.KeyData.SaltValue = s;

		rnd.GetBytes(s);
		encr.SaltValue = s;

		encr.KeyValue = new byte[encr.KeyBits / 8];
		rnd.GetBytes(encr.KeyValue);

		//Get the password key.
		var hashProvider = GetHashProvider(encryptionInfo.KeyEncryptors[0]);
		var baseHash = GetPasswordHash(hashProvider, encr.SaltValue, encryption.Password, encr.SpinCount, encr.HashSize);
		var hashFinal = GetFinalHash(hashProvider, _blockKey_KeyValue, baseHash);
		hashFinal = FixHashSize(hashFinal, encr.KeyBits / 8);

		var encrData = EncryptDataAgile(package, encryptionInfo, hashProvider);

		/**** Data Integrity ****/
		var saltHMAC = new byte[64];
		rnd.GetBytes(saltHMAC);

		SetHMAC(encryptionInfo, hashProvider, saltHMAC, encrData);

		/**** Verifier ****/
		encr.VerifierHashInput = new byte[16];
		rnd.GetBytes(encr.VerifierHashInput);

		encr.VerifierHash = hashProvider.ComputeHash(encr.VerifierHashInput);

		var VerifierInputKey = GetFinalHash(hashProvider, _blockKey_HashInput, baseHash);
		var VerifierHashKey = GetFinalHash(hashProvider, _blockKey_HashValue, baseHash);
		var KeyValueKey = GetFinalHash(hashProvider, _blockKey_KeyValue, baseHash);

		var ms = new MemoryStream();
		EncryptAgileFromKey(encr, VerifierInputKey, encr.VerifierHashInput, 0, encr.VerifierHashInput.Length, encr.SaltValue, ms);
		encr.EncryptedVerifierHashInput = ms.ToArray();

		ms = new MemoryStream();
		EncryptAgileFromKey(encr, VerifierHashKey, encr.VerifierHash, 0, encr.VerifierHash.Length, encr.SaltValue, ms);
		encr.EncryptedVerifierHash = ms.ToArray();

		ms = new MemoryStream();
		EncryptAgileFromKey(encr, KeyValueKey, encr.KeyValue, 0, encr.KeyValue.Length, encr.SaltValue, ms);
		encr.EncryptedKeyValue = ms.ToArray();

		xml = encryptionInfo.Xml.OuterXml;

		var byXml = Encoding.UTF8.GetBytes(xml);

		ms = new MemoryStream();
		ms.Write(BitConverter.GetBytes((ushort)4), 0, 2); //Major Version
		ms.Write(BitConverter.GetBytes((ushort)4), 0, 2); //Minor Version
		ms.Write(BitConverter.GetBytes((uint)0x40), 0, 4); //Reserved
		ms.Write(byXml, 0, byXml.Length);

		var doc = new CompoundDocument();

		//Add the dataspace streams
		CreateDataSpaces(doc);
		//EncryptionInfo...
		doc.Storage.DataStreams.Add("EncryptionInfo", ms.ToArray());
		//...and the encrypted package
		doc.Storage.DataStreams.Add("EncryptedPackage", encrData);

		ms = new MemoryStream();
		doc.Save(ms);
		//ms.Write(e,0,e.Length);
		return ms;
	}

	private byte[] EncryptDataAgile(byte[] data, EncryptionInfoAgile encryptionInfo, HashAlgorithm hashProvider)
	{
		var ke = encryptionInfo.KeyEncryptors[0];
		var aes = Aes.Create();
		aes.KeySize = ke.KeyBits;
		aes.Mode = CipherMode.CBC;
		aes.Padding = PaddingMode.Zeros;

		var pos = 0;
		var segment = 0;

		//Encrypt the data
		var ms = new MemoryStream();
		ms.Write(BitConverter.GetBytes((ulong)data.Length), 0, 8);
		while (pos < data.Length)
		{
			var segmentSize = (int)(data.Length - pos > 4096 ? 4096 : data.Length - pos);

			var ivTmp = new byte[4 + encryptionInfo.KeyData.SaltSize];
			Array.Copy(encryptionInfo.KeyData.SaltValue, 0, ivTmp, 0, encryptionInfo.KeyData.SaltSize);
			Array.Copy(BitConverter.GetBytes(segment), 0, ivTmp, encryptionInfo.KeyData.SaltSize, 4);
			var iv = hashProvider.ComputeHash(ivTmp);

			EncryptAgileFromKey(ke, ke.KeyValue, data, pos, segmentSize, iv, ms);
			pos += segmentSize;
			segment++;
		}

		ms.Flush();
		return ms.ToArray();
	}
	// Set the dataintegrity
	private void SetHMAC(EncryptionInfoAgile ei, HashAlgorithm hashProvider, byte[] salt, byte[] data)
	{
		var iv = GetFinalHash(hashProvider, _blockKey_HmacKey, ei.KeyData.SaltValue);
		var ms = new MemoryStream();
		EncryptAgileFromKey(ei.KeyEncryptors[0], ei.KeyEncryptors[0].KeyValue, salt, 0L, salt.Length, iv, ms);
		ei.DataIntegrity.EncryptedHmacKey = ms.ToArray();

		var h = GetHmacProvider(ei.KeyEncryptors[0], salt);
		var hmacValue = h.ComputeHash(data);

		ms = new MemoryStream();
		iv = GetFinalHash(hashProvider, _blockKey_HmacValue, ei.KeyData.SaltValue);
		EncryptAgileFromKey(ei.KeyEncryptors[0], ei.KeyEncryptors[0].KeyValue, hmacValue, 0L, hmacValue.Length, iv, ms);
		ei.DataIntegrity.EncryptedHmacValue = ms.ToArray();
	}

	private static HMAC GetHmacProvider(EncryptionInfoAgile.EncryptionKeyData ei, byte[] salt) => ei.HashAlgorithm switch
	{
		eHashAlogorithm.MD5 => new HMACMD5(salt),
		eHashAlogorithm.SHA1 => new HMACSHA1(salt),
		eHashAlogorithm.SHA256 => new HMACSHA256(salt),
		eHashAlogorithm.SHA384 => new HMACSHA384(salt),
		eHashAlogorithm.SHA512 => new HMACSHA512(salt),
		_ => throw (new NotSupportedException(string.Format("Hash method {0} not supported.", ei.HashAlgorithm))),
	};

	private MemoryStream EncryptPackageBinary(byte[] package, ExcelEncryption encryption)
	{
		//Create the Encryption Info. This also returns the Encryptionkey
		var encryptionInfo = CreateEncryptionInfo(encryption.Password,
				encryption.Algorithm == EncryptionAlgorithm.AES128 ?
					AlgorithmID.AES128 :
				encryption.Algorithm == EncryptionAlgorithm.AES192 ?
					AlgorithmID.AES192 :
					AlgorithmID.AES256, out var encryptionKey);

		//ILockBytes lb;
		//var iret = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lb);

		//IStorage storage = null;
		//MemoryStream ret = null;

		var doc = new CompoundDocument();
		CreateDataSpaces(doc);

		doc.Storage.DataStreams.Add("EncryptionInfo", encryptionInfo.WriteBinary());

		//Encrypt the package
		var encryptedPackage = EncryptData(encryptionKey, package, false);
		MemoryStream ms = new();
		ms.Write(BitConverter.GetBytes((ulong)package.Length), 0, 8);
		ms.Write(encryptedPackage, 0, encryptedPackage.Length);
		doc.Storage.DataStreams.Add("EncryptedPackage", ms.ToArray());

		var ret = new MemoryStream();
		doc.Save(ret);

		return ret;
	}
	#region "Dataspaces Stream methods"
	private static void CreateDataSpaces(CompoundDocument doc)
	{
		var ds = new CompoundDocument.StoragePart();
		doc.Storage.SubStorage.Add("\x06" + "DataSpaces", ds);
		var ver = new CompoundDocument.StoragePart();
		ds.DataStreams.Add("Version", CreateVersionStream());
		ds.DataStreams.Add("DataSpaceMap", CreateDataSpaceMap());

		var dsInfo = new CompoundDocument.StoragePart();
		ds.SubStorage.Add("DataSpaceInfo", dsInfo);
		dsInfo.DataStreams.Add("StrongEncryptionDataSpace", CreateStrongEncryptionDataSpaceStream());

		var transInfo = new CompoundDocument.StoragePart();
		ds.SubStorage.Add("TransformInfo", transInfo);

		var strEncTrans = new CompoundDocument.StoragePart();
		transInfo.SubStorage.Add("StrongEncryptionTransform", strEncTrans);

		strEncTrans.DataStreams.Add("\x06Primary", CreateTransformInfoPrimary());
	}
	private static byte[] CreateStrongEncryptionDataSpaceStream()
	{
		MemoryStream ms = new();
		BinaryWriter bw = new(ms);

		bw.Write((int)8);       //HeaderLength
		bw.Write((int)1);       //EntryCount

		var tr = "StrongEncryptionTransform";
		bw.Write((int)tr.Length * 2);
		bw.Write(UTF8Encoding.Unicode.GetBytes(tr + "\0")); // end \0 is for padding

		bw.Flush();
		return ms.ToArray();
	}
	private static byte[] CreateVersionStream()
	{
		MemoryStream ms = new();
		BinaryWriter bw = new(ms);

		bw.Write((short)0x3C);  //Major
		bw.Write((short)0);     //Minor
		bw.Write(UTF8Encoding.Unicode.GetBytes("Microsoft.Container.DataSpaces"));
		bw.Write((int)1);       //ReaderVersion
		bw.Write((int)1);       //UpdaterVersion
		bw.Write((int)1);       //WriterVersion

		bw.Flush();
		return ms.ToArray();
	}
	private static byte[] CreateDataSpaceMap()
	{
		MemoryStream ms = new();
		BinaryWriter bw = new(ms);

		bw.Write((int)8);       //HeaderLength
		bw.Write((int)1);       //EntryCount
		var s1 = "EncryptedPackage";
		var s2 = "StrongEncryptionDataSpace";
		bw.Write((int)(s1.Length + s2.Length) * 2 + 0x16);
		bw.Write((int)1);       //ReferenceComponentCount
		bw.Write((int)0);       //Stream=0
		bw.Write((int)s1.Length * 2); //Length s1
		bw.Write(UTF8Encoding.Unicode.GetBytes(s1));
		bw.Write((int)(s2.Length * 2));   //Length s2
		bw.Write(UTF8Encoding.Unicode.GetBytes(s2 + "\0"));   // end \0 is for padding

		bw.Flush();
		return ms.ToArray();
	}
	private static byte[] CreateTransformInfoPrimary()
	{
		MemoryStream ms = new();
		BinaryWriter bw = new(ms);
		var TransformID = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}";
		var TransformName = "Microsoft.Container.EncryptionTransform";
		bw.Write(TransformID.Length * 2 + 12);
		bw.Write((int)1);
		bw.Write(TransformID.Length * 2);
		bw.Write(UTF8Encoding.Unicode.GetBytes(TransformID));
		bw.Write(TransformName.Length * 2);
		bw.Write(UTF8Encoding.Unicode.GetBytes(TransformName + "\0"));
		bw.Write((int)1);   //ReaderVersion
		bw.Write((int)1);   //UpdaterVersion
		bw.Write((int)1);   //WriterVersion

		bw.Write((int)0);
		bw.Write((int)0);
		bw.Write((int)0);       //CipherMode
		bw.Write((int)4);       //Reserved

		bw.Flush();
		return ms.ToArray();
	}
	#endregion
	/// <summary>
	/// Create an EncryptionInfo object to encrypt a workbook
	/// </summary>
	/// <param name="password">The password</param>
	/// <param name="algID"></param>
	/// <param name="key">The Encryption key</param>
	/// <returns></returns>
	private EncryptionInfoBinary CreateEncryptionInfo(string password, AlgorithmID algID, out byte[] key)
	{
		if (algID is AlgorithmID.Flags or AlgorithmID.RC4)
		{
			throw (new ArgumentException("algID must be AES128, AES192 or AES256"));
		}

		var encryptionInfo = new EncryptionInfoBinary
		{
			MajorVersion = 4,
			MinorVersion = 2,
			Flags = Flags.fAES | Flags.fCryptoAPI
		};

		//Header
		encryptionInfo.Header = new EncryptionHeader
		{
			AlgID = algID,
			AlgIDHash = AlgorithmHashID.SHA1,
			Flags = encryptionInfo.Flags,
			KeySize =
			(algID == AlgorithmID.AES128 ? 0x80 : algID == AlgorithmID.AES192 ? 0xC0 : 0x100),
			ProviderType = ProviderType.AES,
			CSPName = "Microsoft Enhanced RSA and AES Cryptographic Provider\0",
			Reserved1 = 0,
			Reserved2 = 0,
			SizeExtra = 0
		};

		//Verifier
		encryptionInfo.Verifier = new EncryptionVerifier
		{
			Salt = new byte[16]
		};

		var rnd = RandomNumberGenerator.Create();
		rnd.GetBytes(encryptionInfo.Verifier.Salt);
		encryptionInfo.Verifier.SaltSize = 0x10;

		key = GetPasswordHashBinary(password, encryptionInfo);

		var verifier = new byte[16];
		rnd.GetBytes(verifier);
		encryptionInfo.Verifier.EncryptedVerifier = EncryptData(key, verifier, true);

		//AES = 32 Bits
		encryptionInfo.Verifier.VerifierHashSize = 0x20;
		var verifierHash = SHA1.HashData(verifier);

		encryptionInfo.Verifier.EncryptedVerifierHash = EncryptData(key, verifierHash, false);

		return encryptionInfo;
	}
	private static byte[] EncryptData(byte[] key, byte[] data, bool useDataSize)
	{
		var aes = Aes.Create();
		aes.KeySize = key.Length * 8;
		aes.Mode = CipherMode.ECB;
		aes.Padding = PaddingMode.Zeros;

		//Encrypt the data
		var crypt = aes.CreateEncryptor(key, null);
		var ms = new MemoryStream();
		var cs = new CryptoStream(ms, crypt, CryptoStreamMode.Write);
		cs.Write(data, 0, data.Length);

		cs.FlushFinalBlock();

		byte[] ret;
		if (useDataSize)
		{
			ret = new byte[data.Length];
			ms.Seek(0, SeekOrigin.Begin);
			ms.Read(ret, 0, data.Length);  //Truncate any padded Zeros
			return ret;
		}
		else
		{
			return ms.ToArray();
		}
	}
	private MemoryStream GetStreamFromPackage(CompoundDocument doc, ExcelEncryption encryption)
	{
		var ret = new MemoryStream();
		if (doc.Storage.DataStreams.ContainsKey("EncryptionInfo") ||
		   doc.Storage.DataStreams.ContainsKey("EncryptedPackage"))
		{
			var encryptionInfo = EncryptionInfo.ReadBinary(doc.Storage.DataStreams["EncryptionInfo"]);

			return DecryptDocument(doc.Storage.DataStreams["EncryptedPackage"], encryptionInfo, encryption.Password);
		}
		else
		{
			throw (new InvalidDataException("Invalid document. EncryptionInfo or EncryptedPackage stream is missing"));
		}
	}

	/// <summary>
	/// Decrypt a document
	/// </summary>
	/// <param name="data">The Encrypted data</param>
	/// <param name="encryptionInfo">Encryption Info object</param>
	/// <param name="password">The password</param>
	/// <returns></returns>
	private MemoryStream DecryptDocument(byte[] data, EncryptionInfo encryptionInfo, string password)
	{
		var size = BitConverter.ToInt64(data, 0);

		var encryptedData = new byte[data.Length - 8];
		Array.Copy(data, 8, encryptedData, 0, encryptedData.Length);

		return encryptionInfo is EncryptionInfoBinary binary
			? DecryptBinary(binary, password, size, encryptedData)
			: DecryptAgile((EncryptionInfoAgile)encryptionInfo, password, size, encryptedData, data);
	}

	readonly byte[] _blockKey_HashInput = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
	readonly byte[] _blockKey_HashValue = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
	readonly byte[] _blockKey_KeyValue = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
	readonly byte[] _blockKey_HmacKey = [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];//MSOFFCRYPTO 2.3.4.14 section 3
	readonly byte[] _blockKey_HmacValue = [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];//MSOFFCRYPTO 2.3.4.14 section 5

	private MemoryStream DecryptAgile(EncryptionInfoAgile encryptionInfo, string password, long size, byte[] encryptedData, byte[] data)
	{
		MemoryStream doc = new();

		if (encryptionInfo.KeyData.CipherAlgorithm == eCipherAlgorithm.AES)
		{
			var encr = encryptionInfo.KeyEncryptors[0];
			var hashProvider = GetHashProvider(encr);
			var hashProviderDataKey = GetHashProvider(encryptionInfo.KeyData);

			var baseHash = GetPasswordHash(hashProvider, encr.SaltValue, password, encr.SpinCount, encr.HashSize);

			//Get the keys for the verifiers and the key value
			var valInputKey = GetFinalHash(hashProvider, _blockKey_HashInput, baseHash);
			var valHashKey = GetFinalHash(hashProvider, _blockKey_HashValue, baseHash);
			var valKeySizeKey = GetFinalHash(hashProvider, _blockKey_KeyValue, baseHash);

			//Decrypt
			encr.VerifierHashInput = DecryptAgileFromKey(encr, valInputKey, encr.EncryptedVerifierHashInput, encr.SaltSize, encr.SaltValue);
			encr.VerifierHash = DecryptAgileFromKey(encr, valHashKey, encr.EncryptedVerifierHash, encr.HashSize, encr.SaltValue);
			encr.KeyValue = DecryptAgileFromKey(encr, valKeySizeKey, encr.EncryptedKeyValue, encryptionInfo.KeyData.KeyBits / 8, encr.SaltValue);

			if (IsPasswordValid(hashProvider, encr))
			{
				var ivhmac = GetFinalHash(hashProviderDataKey, _blockKey_HmacKey, encryptionInfo.KeyData.SaltValue);
				var key = DecryptAgileFromKey(encryptionInfo.KeyData, encr.KeyValue, encryptionInfo.DataIntegrity.EncryptedHmacKey, encryptionInfo.KeyData.HashSize, ivhmac);

				ivhmac = GetFinalHash(hashProviderDataKey, _blockKey_HmacValue, encryptionInfo.KeyData.SaltValue);
				var value = DecryptAgileFromKey(encryptionInfo.KeyData, encr.KeyValue, encryptionInfo.DataIntegrity.EncryptedHmacValue, encryptionInfo.KeyData.HashSize, ivhmac);

				var hmca = GetHmacProvider(encryptionInfo.KeyData, key);
				var v2 = hmca.ComputeHash(data);

				for (var i = 0; i < v2.Length; i++)
				{
					if (value[i] != v2[i])
					{
						throw (new Exception("Dataintegrity key mismatch"));
					}
				}

				var pos = 0;
				uint segment = 0;
				while (pos < size)
				{
					var segmentSize = (int)(size - pos > 4096 ? 4096 : size - pos);
					var bufferSize = (int)(encryptedData.Length - pos > 4096 ? 4096 : encryptedData.Length - pos);
					var ivTmp = new byte[4 + encryptionInfo.KeyData.SaltSize];
					Array.Copy(encryptionInfo.KeyData.SaltValue, 0, ivTmp, 0, encryptionInfo.KeyData.SaltSize);
					Array.Copy(BitConverter.GetBytes(segment), 0, ivTmp, encryptionInfo.KeyData.SaltSize, 4);
					var iv = hashProviderDataKey.ComputeHash(ivTmp);
					var buffer = new byte[bufferSize];
					Array.Copy(encryptedData, pos, buffer, 0, bufferSize);

					var b = DecryptAgileFromKey(encryptionInfo.KeyData, encr.KeyValue, buffer, segmentSize, iv);
					doc.Write(b, 0, b.Length);
					pos += segmentSize;
					segment++;
				}

				doc.Flush();
				return doc;
			}
			else
			{
				throw (new SecurityException("Invalid password"));
			}
		}

		return null;
	}
	private static HashAlgorithm GetHashProvider(EncryptionInfoAgile.EncryptionKeyData encr) => encr.HashAlgorithm switch
	{
		eHashAlogorithm.MD5 => MD5.Create(),
		//case eHashAlogorithm.RIPEMD160:
		//    return new RIPEMD160Managed();                    
		eHashAlogorithm.SHA1 => SHA1.Create(),
		eHashAlogorithm.SHA256 => SHA256.Create(),
		eHashAlogorithm.SHA384 => SHA384.Create(),
		eHashAlogorithm.SHA512 => SHA512.Create(),
		_ => throw new NotSupportedException(string.Format("Hash provider is unsupported. {0}", encr.HashAlgorithm)),
	};
	private MemoryStream DecryptBinary(EncryptionInfoBinary encryptionInfo, string password, long size, byte[] encryptedData)
	{
		MemoryStream doc = new();

		if (encryptionInfo.Header.AlgID == AlgorithmID.AES128 || (encryptionInfo.Header.AlgID == AlgorithmID.Flags && ((encryptionInfo.Flags & (Flags.fAES | Flags.fExternal | Flags.fCryptoAPI)) == (Flags.fAES | Flags.fCryptoAPI)))
			||
			encryptionInfo.Header.AlgID == AlgorithmID.AES192
			||
			encryptionInfo.Header.AlgID == AlgorithmID.AES256
			)
		{
			var decryptKey = Aes.Create();
			decryptKey.KeySize = encryptionInfo.Header.KeySize;
			decryptKey.Mode = CipherMode.ECB;
			decryptKey.Padding = PaddingMode.None;

			var key = GetPasswordHashBinary(password, encryptionInfo);
			if (IsPasswordValid(key, encryptionInfo))
			{
				var decryptor = decryptKey.CreateDecryptor(
														 key,
														 null);

				var dataStream = new MemoryStream(encryptedData);
				var cryptoStream = new CryptoStream(dataStream,
															  decryptor,
															  CryptoStreamMode.Read);

				var decryptedData = new byte[size];

				cryptoStream.Read(decryptedData, 0, (int)size);
				doc.Write(decryptedData, 0, (int)size);
			}
			else
			{
				throw (new UnauthorizedAccessException("Invalid password"));
			}
		}

		return doc;
	}
	/// <summary>
	/// Validate the password
	/// </summary>
	/// <param name="key">The encryption key</param>
	/// <param name="encryptionInfo">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
	/// <returns></returns>
	private static bool IsPasswordValid(byte[] key, EncryptionInfoBinary encryptionInfo)
	{
		var decryptKey = Aes.Create();
		decryptKey.KeySize = encryptionInfo.Header.KeySize;
		decryptKey.Mode = CipherMode.ECB;
		decryptKey.Padding = PaddingMode.None;

		var decryptor = decryptKey.CreateDecryptor(
												 key,
												 null);


		//Decrypt the verifier
		MemoryStream dataStream = new(encryptionInfo.Verifier.EncryptedVerifier);
		CryptoStream cryptoStream = new(dataStream,
													  decryptor,
													  CryptoStreamMode.Read);
		var decryptedVerifier = new byte[16];
		cryptoStream.Read(decryptedVerifier, 0, 16);

		dataStream = new MemoryStream(encryptionInfo.Verifier.EncryptedVerifierHash);

		cryptoStream = new CryptoStream(dataStream,
											decryptor,
											CryptoStreamMode.Read);

		//Decrypt the verifier hash
		var decryptedVerifierHash = new byte[16];
		cryptoStream.Read(decryptedVerifierHash, 0, (int)16);

		//Get the hash for the decrypted verifier
		var hash = SHA1.HashData(decryptedVerifier);

		//Equal?
		for (var i = 0; i < 16; i++)
		{
			if (hash[i] != decryptedVerifierHash[i])
			{
				return false;
			}
		}

		return true;
	}
	/// <summary>
	/// Validate the password
	/// </summary>
	/// <param name="sha">The hash algorithm</param>
	/// <param name="encr">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
	/// <returns></returns>
	private static bool IsPasswordValid(HashAlgorithm sha, EncryptionInfoAgile.EncryptionKeyEncryptor encr)
	{
		var valHash = sha.ComputeHash(encr.VerifierHashInput);

		//Equal?
		for (var i = 0; i < valHash.Length; i++)
		{
			if (encr.VerifierHash[i] != valHash[i])
			{
				return false;
			}
		}

		return true;
	}

	private static byte[] DecryptAgileFromKey(EncryptionInfoAgile.EncryptionKeyData encr, byte[] key, byte[] encryptedData, long size, byte[] iv)
	{
		var decryptKey = GetEncryptionAlgorithm(encr);
		decryptKey.BlockSize = encr.BlockSize << 3;
		decryptKey.KeySize = encr.KeyBits;
		decryptKey.Mode = CipherMode.CBC;
		decryptKey.Padding = PaddingMode.Zeros;

		var decryptor = decryptKey.CreateDecryptor(
													FixHashSize(key, encr.KeyBits / 8),
													FixHashSize(iv, encr.BlockSize, 0x36));


		MemoryStream dataStream = new(encryptedData);

		CryptoStream cryptoStream = new(dataStream,
														decryptor,
														CryptoStreamMode.Read);

		var decryptedData = new byte[size];

		cryptoStream.Read(decryptedData, 0, (int)size);
		return decryptedData;
	}

	private static SymmetricAlgorithm GetEncryptionAlgorithm(EncryptionInfoAgile.EncryptionKeyData encr) => encr.CipherAlgorithm switch
	{
		eCipherAlgorithm.AES => Aes.Create(),
		//case eCipherAlgorithm.DES:
		//    return new DESCryptoServiceProvider();
		eCipherAlgorithm.TRIPLE_DES or eCipherAlgorithm.TRIPLE_DES_112 => TripleDES.Create(),
		//case eCipherAlgorithm.RC2:
		//    return new RC2CryptoServiceProvider();                    
		_ => throw (new NotSupportedException(string.Format("Unsupported Cipher Algorithm: {0}", encr.CipherAlgorithm.ToString()))),
	};

	private static void EncryptAgileFromKey(EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] key, byte[] data, long pos, long size, byte[] iv, MemoryStream ms)
	{
		var encryptKey = GetEncryptionAlgorithm(encr);
		encryptKey.BlockSize = encr.BlockSize << 3;
		encryptKey.KeySize = encr.KeyBits;
		encryptKey.Mode = CipherMode.CBC;
		encryptKey.Padding = PaddingMode.Zeros;

		var encryptor = encryptKey.CreateEncryptor(
													FixHashSize(key, encr.KeyBits / 8),
													FixHashSize(iv, 16, 0x36));


		CryptoStream cryptoStream = new(ms,
													 encryptor,
													 CryptoStreamMode.Write);

		var cryptoSize = size % encr.BlockSize == 0 ? size : (size + (encr.BlockSize - (size % encr.BlockSize)));
		var buffer = new byte[size];
		Array.Copy(data, (int)pos, buffer, 0, (int)size);
		cryptoStream.Write(buffer, 0, (int)size);
		while (size % encr.BlockSize != 0)
		{
			cryptoStream.WriteByte(0);
			size++;
		}
	}

	/// <summary>
	/// Create the hash.
	/// This method is written with the help of Lyquidity library, many thanks for this nice sample
	/// </summary>
	/// <param name="password">The password</param>
	/// <param name="encryptionInfo">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
	/// <returns>The hash to encrypt the document</returns>
	private byte[] GetPasswordHashBinary(string password, EncryptionInfoBinary encryptionInfo)
	{
		byte[] hash = null;
		var tempHash = new byte[4 + 20];    //Iterator + prev. hash
		try
		{
			HashAlgorithm hashProvider;
			if (encryptionInfo.Header.AlgIDHash == AlgorithmHashID.SHA1 || encryptionInfo.Header.AlgIDHash == AlgorithmHashID.App && (encryptionInfo.Flags & Flags.fExternal) == 0)
			{
				hashProvider = SHA1.Create();
			}
			else if (encryptionInfo.Header.KeySize is > 0 and < 80)
			{
				throw new NotSupportedException("RC4 Hash provider is not supported. Must be SHA1(AlgIDHash == 0x8004)");
			}
			else
			{
				throw new NotSupportedException("Hash provider is invalid. Must be SHA1(AlgIDHash == 0x8004)");
			}

			hash = GetPasswordHash(hashProvider, encryptionInfo.Verifier.Salt, password, 50000, 20);

			// Append "block" (0)
			Array.Copy(hash, tempHash, hash.Length);
			Array.Copy(System.BitConverter.GetBytes(0), 0, tempHash, hash.Length, 4);
			hash = hashProvider.ComputeHash(tempHash);

			/***** Now use the derived key algorithm *****/
			var derivedKey = new byte[64];
			var keySizeBytes = encryptionInfo.Header.KeySize / 8;

			//First XOR hash bytes with 0x36 and fill the rest with 0x36
			for (var i = 0; i < derivedKey.Length; i++)
				derivedKey[i] = (byte)(i < hash.Length ? 0x36 ^ hash[i] : 0x36);


			var X1 = hashProvider.ComputeHash(derivedKey);

			//if verifier size is bigger than the key size we can return X1
			if ((int)encryptionInfo.Verifier.VerifierHashSize > keySizeBytes)
				return FixHashSize(X1, keySizeBytes);

			//Else XOR hash bytes with 0x5C and fill the rest with 0x5C
			for (var i = 0; i < derivedKey.Length; i++)
				derivedKey[i] = (byte)(i < hash.Length ? 0x5C ^ hash[i] : 0x5C);

			var X2 = hashProvider.ComputeHash(derivedKey);

			//Join the two and return 
			var join = new byte[X1.Length + X2.Length];

			Array.Copy(X1, 0, join, 0, X1.Length);
			Array.Copy(X2, 0, join, X1.Length, X2.Length);


			return FixHashSize(join, keySizeBytes);
		}
		catch (Exception ex)
		{
			throw (new Exception("An error occured when the encryptionkey was created", ex));
		}
	}

	/// <summary>
	/// Create the hash.
	/// This method is written with the help of Lyquidity library, many thanks for this nice sample
	/// </summary>
	/// <param name="password">The password</param>
	/// <param name="encr">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
	/// <param name="blockKey">The block key appended to the hash to obtain the final hash</param>
	/// <returns>The hash to encrypt the document</returns>
	private byte[] GetPasswordHashAgile(string password, EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] blockKey)
	{
		try
		{
			var hashProvider = GetHashProvider(encr);
			var hash = GetPasswordHash(hashProvider, encr.SaltValue, password, encr.SpinCount, encr.HashSize);
			var hashFinal = GetFinalHash(hashProvider, blockKey, hash);

			return FixHashSize(hashFinal, encr.KeyBits / 8);
		}
		catch (Exception ex)
		{
			throw (new Exception("An error occured when the encryptionkey was created", ex));
		}
	}
	private static byte[] GetFinalHash(HashAlgorithm hashProvider, byte[] blockKey, byte[] hash)
	{
		//2.3.4.13 MS-OFFCRYPTO
		var tempHash = new byte[hash.Length + blockKey.Length];
		Array.Copy(hash, tempHash, hash.Length);
		Array.Copy(blockKey, 0, tempHash, hash.Length, blockKey.Length);
		var hashFinal = hashProvider.ComputeHash(tempHash);
		return hashFinal;
	}
	private static byte[] GetPasswordHash(HashAlgorithm hashProvider, byte[] salt, string password, int spinCount, int hashSize)
	{
		byte[] hash = null;
		var tempHash = new byte[4 + hashSize];    //Iterator + prev. hash
		hash = hashProvider.ComputeHash(CombinePassword(salt, password));

		//Iterate "spinCount" times, inserting i in first 4 bytes and then the prev. hash in byte 5-24
		for (var i = 0; i < spinCount; i++)
		{
			Array.Copy(BitConverter.GetBytes(i), tempHash, 4);
			Array.Copy(hash, 0, tempHash, 4, hash.Length);

			hash = hashProvider.ComputeHash(tempHash);
		}

		return hash;
	}
	private static byte[] FixHashSize(byte[] hash, int size, byte fill = 0)
	{
		if (hash.Length == size)
			return hash;
		else if (hash.Length < size)
		{
			var buff = new byte[size];
			Array.Copy(hash, buff, hash.Length);
			for (var i = hash.Length; i < size; i++)
			{
				buff[i] = fill;
			}

			return buff;
		}
		else
		{
			var buff = new byte[size];
			Array.Copy(hash, buff, size);
			return buff;
		}
	}
	private static byte[] CombinePassword(byte[] salt, string password)
	{
		if (password == "")
		{
			password = "VelvetSweatshop";   //Used if Password is blank
		}
		// Convert password to unicode...
		var passwordBuf = UnicodeEncoding.Unicode.GetBytes(password);

		var inputBuf = new byte[salt.Length + passwordBuf.Length];
		Array.Copy(salt, inputBuf, salt.Length);
		Array.Copy(passwordBuf, 0, inputBuf, salt.Length, passwordBuf.Length);
		return inputBuf;
	}
	internal static ushort CalculatePasswordHash(string Password)
	{
		//Calculate the hash
		//Thanks to Kohei Yoshida for the sample http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/
		ushort hash = 0;
		for (var i = Password.Length - 1; i >= 0; i--)
		{
			hash ^= Password[i];
			hash = (ushort)(((ushort)((hash >> 14) & 0x01))
							|
							((ushort)((hash << 1) & 0x7FFF)));  //Shift 1 to the left. Overflowing bit 15 goes into bit 0
		}

		hash ^= (0x8000 | ('N' << 8) | 'K'); //Xor NK with high bit set(0xCE4B)
		hash ^= (ushort)Password.Length;

		return hash;
	}
}
