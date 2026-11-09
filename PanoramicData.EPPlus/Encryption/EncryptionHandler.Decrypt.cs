using OfficeOpenXml.Utils.CompundDocument;
using System.Security;
using System.Security.Cryptography;

namespace OfficeOpenXml.Encryption;

internal partial class EncryptedPackageHandler
{
	private MemoryStream GetStreamFromPackage(CompoundDocument doc, ExcelEncryption encryption)
	{
		if(doc.Storage.DataStreams.ContainsKey("EncryptionInfo") ||
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
				VerifyDataIntegrity(encryptionInfo, encr, hashProviderDataKey, data);
				DecryptAgileSegments(encryptionInfo, encr, hashProviderDataKey, size, encryptedData, doc);
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

	private void VerifyDataIntegrity(EncryptionInfoAgile encryptionInfo, EncryptionInfoAgile.EncryptionKeyEncryptor encr, HashAlgorithm hashProviderDataKey, byte[] data)
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
	}

	private static void DecryptAgileSegments(EncryptionInfoAgile encryptionInfo, EncryptionInfoAgile.EncryptionKeyEncryptor encr, HashAlgorithm hashProviderDataKey, long size, byte[] encryptedData, MemoryStream doc)
	{
		var pos = 0;
		uint segment = 0;
		while (pos < size)
		{
			var segmentSize = (int)(size - pos > 4096 ? 4096 : size - pos);
			var bufferSize = encryptedData.Length - pos > 4096 ? 4096 : encryptedData.Length - pos;
			var iv = GetSegmentIv(hashProviderDataKey, encryptionInfo.KeyData, BitConverter.GetBytes(segment));
			var buffer = new byte[bufferSize];
			Array.Copy(encryptedData, pos, buffer, 0, bufferSize);

			var b = DecryptAgileFromKey(encryptionInfo.KeyData, encr.KeyValue, buffer, segmentSize, iv);
			doc.Write(b, 0, b.Length);
			pos += segmentSize;
			segment++;
		}
	}
	private static HashAlgorithm GetHashProvider(EncryptionInfoAgile.EncryptionKeyData encr) => encr.HashAlgorithm switch
	{
		eHashAlogorithm.MD5 => MD5.Create(),
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
			decryptKey.Mode = CipherMode.ECB; // nosemgrep: ECB is mandated by MS-OFFCRYPTO standard encryption
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

				cryptoStream.ReadAtLeast(decryptedData.AsSpan(0, (int)size), (int)size, throwOnEndOfStream: false);
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
		cryptoStream.ReadAtLeast(decryptedVerifier.AsSpan(0, 16), 16, throwOnEndOfStream: false);

		dataStream = new MemoryStream(encryptionInfo.Verifier.EncryptedVerifierHash);

		cryptoStream = new CryptoStream(dataStream,
											decryptor,
											CryptoStreamMode.Read);

		//Decrypt the verifier hash
		var decryptedVerifierHash = new byte[16];
		cryptoStream.ReadAtLeast(decryptedVerifierHash.AsSpan(0, 16), 16, throwOnEndOfStream: false);

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
		decryptKey.Mode = CipherMode.CBC; // nosemgrep: CBC is mandated by MS-OFFCRYPTO agile encryption
		decryptKey.Padding = PaddingMode.Zeros;

		var decryptor = decryptKey.CreateDecryptor(
													FixHashSize(key, encr.KeyBits / 8),
													FixHashSize(iv, encr.BlockSize, 0x36));


		MemoryStream dataStream = new(encryptedData);

		CryptoStream cryptoStream = new(dataStream,
														decryptor,
														CryptoStreamMode.Read);

		var decryptedData = new byte[size];

		cryptoStream.ReadAtLeast(decryptedData.AsSpan(0, (int)size), (int)size, throwOnEndOfStream: false);
		return decryptedData;
	}
}
