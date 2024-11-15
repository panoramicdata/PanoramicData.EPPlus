// Shared.cs
// ------------------------------------------------------------------
//
// Copyright (c) 2006-2011 Dino Chiesa.
// All rights reserved.
//
// This code module is part of DotNetZip, a zipfile class library.
//
// ------------------------------------------------------------------
//
// This code is licensed under the Microsoft Public License.
// See the file License.txt for the license details.
// More info on: http://dotnetzip.codeplex.com
//
// ------------------------------------------------------------------
//
// Last Saved: <2011-August-02 19:41:01>
//
// ------------------------------------------------------------------
//
// This module defines some shared utility classes and methods.
//
// Created: Tue, 27 Mar 2007  15:30
//

using System;
using System.IO;

namespace OfficeOpenXml.Packaging.DotNetZip;

/// <summary>
/// Collects general purpose utility methods.
/// </summary>
internal static class SharedUtilities
{
	/// private null constructor
	//private SharedUtilities() { }

	// workitem 8423
	public static long GetFileLength(string fileName)
	{
		if (!File.Exists(fileName))
			throw new FileNotFoundException(fileName);

		var fileLength = 0L;
		var fs = FileShare.ReadWrite;
		// FileShare.Delete is not defined for the Compact Framework
		fs |= FileShare.Delete;
		using (var s = File.Open(fileName, FileMode.Open, FileAccess.Read, fs))
		{
			fileLength = s.Length;
		}

		return fileLength;
	}


	[System.Diagnostics.Conditional("NETCF")]
	public static void Workaround_Ladybug318918(Stream s) =>
		// This is a workaround for this issue:
		// https://connect.microsoft.com/VisualStudio/feedback/details/318918
		// It's required only on NETCF.
		s.Flush();

	private static readonly System.Text.RegularExpressions.Regex doubleDotRegex1 =
		new(@"^(.*/)?([^/\\.]+/\\.\\./)(.+)$");

	private static string SimplifyFwdSlashPath(string path)
	{
		if (path.StartsWith("./")) path = path[2..];
		path = path.Replace("/./", "/");

		// Replace foo/anything/../bar with foo/bar
		path = doubleDotRegex1.Replace(path, "$1$3");
		return path;
	}


	/// <summary>
	/// Utility routine for transforming path names from filesystem format (on Windows that means backslashes) to
	/// a format suitable for use within zipfiles. This means trimming the volume letter and colon (if any) And
	/// swapping backslashes for forward slashes.
	/// </summary>
	/// <param name="pathName">source path.</param>
	/// <returns>transformed path</returns>
	public static string NormalizePathForUseInZipFile(string pathName)
	{
		// boundary case
		if (string.IsNullOrEmpty(pathName)) return pathName;

		// trim volume if necessary
		if (pathName.Length >= 2 && pathName[1] == ':' && pathName[2] == '\\')
			pathName = pathName[3..];

		// swap slashes
		pathName = pathName.Replace('\\', '/');

		// trim all leading slashes
		while (pathName.StartsWith("/")) pathName = pathName[1..];

		return SimplifyFwdSlashPath(pathName);
	}


	static readonly System.Text.Encoding ibm437 = System.Text.Encoding.GetEncoding("UTF-8");
	static readonly System.Text.Encoding utf8 = System.Text.Encoding.GetEncoding("UTF-8");

	internal static byte[] StringToByteArray(string value, System.Text.Encoding encoding)
	{
		var a = encoding.GetBytes(value);
		return a;
	}
	internal static byte[] StringToByteArray(string value) => StringToByteArray(value, ibm437);

	//internal static byte[] Utf8StringToByteArray(string value)
	//{
	//    return StringToByteArray(value, utf8);
	//}

	//internal static string StringFromBuffer(byte[] buf, int maxlength)
	//{
	//    return StringFromBuffer(buf, maxlength, ibm437);
	//}

	internal static string Utf8StringFromBuffer(byte[] buf) => StringFromBuffer(buf, utf8);

	internal static string StringFromBuffer(byte[] buf, System.Text.Encoding encoding)
	{
		// this form of the GetString() method is required for .NET CF compatibility
		var s = encoding.GetString(buf, 0, buf.Length);
		return s;
	}


	internal static int ReadSignature(Stream s)
	{
		var x = 0;
		try { x = _ReadFourBytes(s, "n/a"); }
		catch (BadReadException) { }

		return x;
	}


	internal static int ReadEntrySignature(Stream s)
	{
		// handle the case of ill-formatted zip archives - includes a data descriptor
		// when none is expected.
		var x = 0;
		try
		{
			x = _ReadFourBytes(s, "n/a");
			if (x == ZipConstants.ZipEntryDataDescriptorSignature)
			{
				// advance past data descriptor - 12 bytes if not zip64
				s.Seek(12, SeekOrigin.Current);
				// workitem 10178
				Workaround_Ladybug318918(s);
				x = _ReadFourBytes(s, "n/a");
				if (x != ZipConstants.ZipEntrySignature)
				{
					// Maybe zip64 was in use for the prior entry.
					// Therefore, skip another 8 bytes.
					s.Seek(8, SeekOrigin.Current);
					// workitem 10178
					Workaround_Ladybug318918(s);
					x = _ReadFourBytes(s, "n/a");
					if (x != ZipConstants.ZipEntrySignature)
					{
						// seek back to the first spot
						s.Seek(-24, SeekOrigin.Current);
						// workitem 10178
						Workaround_Ladybug318918(s);
						x = _ReadFourBytes(s, "n/a");
					}
				}
			}
		}
		catch (BadReadException) { }

		return x;
	}


	internal static int ReadInt(Stream s) => _ReadFourBytes(s, "Could not read block - no data!  (position 0x{0:X8})");

	private static int _ReadFourBytes(Stream s, string message)
	{
		var n = 0;
		var block = new byte[4];

		n = s.Read(block, 0, block.Length);
		if (n != block.Length) throw new BadReadException(string.Format(message, s.Position));
		var data = unchecked(((block[3] * 256 + block[2]) * 256 + block[1]) * 256 + block[0]);
		return data;
	}



	/// <summary>
	///   Finds a signature in the zip stream. This is useful for finding
	///   the end of a zip entry, for example, or the beginning of the next ZipEntry.
	/// </summary>
	///
	/// <remarks>
	///   <para>
	///     Scans through 64k at a time.
	///   </para>
	///
	///   <para>
	///     If the method fails to find the requested signature, the stream Position
	///     after completion of this method is unchanged. If the method succeeds in
	///     finding the requested signature, the stream position after completion is
	///     direct AFTER the signature found in the stream.
	///   </para>
	/// </remarks>
	///
	/// <param name="stream">The stream to search</param>
	/// <param name="SignatureToFind">The 4-byte signature to find</param>
	/// <returns>The number of bytes read</returns>
	internal static long FindSignature(Stream stream, int SignatureToFind)
	{
		var startingPosition = stream.Position;

		var BATCH_SIZE = 65536; //  8192;
		var targetBytes = new byte[4];
		targetBytes[0] = (byte)(SignatureToFind >> 24);
		targetBytes[1] = (byte)((SignatureToFind & 0x00FF0000) >> 16);
		targetBytes[2] = (byte)((SignatureToFind & 0x0000FF00) >> 8);
		targetBytes[3] = (byte)(SignatureToFind & 0x000000FF);
		var batch = new byte[BATCH_SIZE];
		var n = 0;
		var success = false;
		do
		{
			n = stream.Read(batch, 0, batch.Length);
			if (n != 0)
			{
				for (var i = 0; i < n; i++)
				{
					if (batch[i] == targetBytes[3])
					{
						var curPosition = stream.Position;
						stream.Seek(i - n, SeekOrigin.Current);
						// workitem 10178
						Workaround_Ladybug318918(stream);

						// workitem 7711
						var sig = ReadSignature(stream);

						success = sig == SignatureToFind;
						if (!success)
						{
							stream.Seek(curPosition, SeekOrigin.Begin);
							// workitem 10178
							Workaround_Ladybug318918(stream);
						}
						else
							break; // out of for loop
					}
				}
			}
			else break;
			if (success) break;

		} while (true);

		if (!success)
		{
			stream.Seek(startingPosition, SeekOrigin.Begin);
			// workitem 10178
			Workaround_Ladybug318918(stream);
			return -1;  // or throw?
		}

		// subtract 4 for the signature.
		var bytesRead = stream.Position - startingPosition - 4;

		return bytesRead;
	}


	// If I have a time in the .NET environment, and I want to use it for
	// SetWastWriteTime() etc, then I need to adjust it for Win32.
	internal static DateTime AdjustTime_Reverse(DateTime time)
	{
		if (time.Kind == DateTimeKind.Utc) return time;
		var adjusted = time;
		if (DateTime.Now.IsDaylightSavingTime() && !time.IsDaylightSavingTime())
			adjusted = time - new TimeSpan(1, 0, 0);

		else if (!DateTime.Now.IsDaylightSavingTime() && time.IsDaylightSavingTime())
			adjusted = time + new TimeSpan(1, 0, 0);

		return adjusted;
	}

#if NECESSARY
        // If I read a time from a file with GetLastWriteTime() (etc), I need
        // to adjust it for display in the .NET environment.
        internal static DateTime AdjustTime_Forward(DateTime time)
        {
            if (time.Kind == DateTimeKind.Utc) return time;
            DateTime adjusted = time;
            if (DateTime.Now.IsDaylightSavingTime() && !time.IsDaylightSavingTime())
                adjusted = time + new System.TimeSpan(1, 0, 0);

            else if (!DateTime.Now.IsDaylightSavingTime() && time.IsDaylightSavingTime())
                adjusted = time - new System.TimeSpan(1, 0, 0);

            return adjusted;
        }
#endif


	internal static DateTime PackedToDateTime(int packedDateTime)
	{
		// workitem 7074 & workitem 7170
		if (packedDateTime is 0xFFFF or 0)
			return new DateTime(1995, 1, 1, 0, 0, 0, 0);  // return a fixed date when none is supplied.

		var packedTime = unchecked((short)(packedDateTime & 0x0000ffff));
		var packedDate = unchecked((short)((packedDateTime & 0xffff0000) >> 16));

		var year = 1980 + ((packedDate & 0xFE00) >> 9);
		var month = (packedDate & 0x01E0) >> 5;
		var day = packedDate & 0x001F;

		var hour = (packedTime & 0xF800) >> 11;
		var minute = (packedTime & 0x07E0) >> 5;
		//int second = packedTime & 0x001F;
		var second = (packedTime & 0x001F) * 2;

		// validation and error checking.
		// this is not foolproof but will catch most errors.
		if (second >= 60) { minute++; second = 0; }

		if (minute >= 60) { hour++; minute = 0; }

		if (hour >= 24) { day++; hour = 0; }

		var d = DateTime.Now;
		var success = false;
		try
		{
			d = new DateTime(year, month, day, hour, minute, second, 0);
			success = true;
		}
		catch (ArgumentOutOfRangeException)
		{
			if (year == 1980 && (month == 0 || day == 0))
			{
				try
				{
					d = new DateTime(1980, 1, 1, hour, minute, second, 0);
					success = true;
				}
				catch (ArgumentOutOfRangeException)
				{
					try
					{
						d = new DateTime(1980, 1, 1, 0, 0, 0, 0);
						success = true;
					}
					catch (ArgumentOutOfRangeException) { }

				}
			}
			// workitem 8814
			// my god, I can't believe how many different ways applications
			// can mess up a simple date format.
			else
			{
				try
				{
					while (year < 1980) year++;
					while (year > 2030) year--;
					while (month < 1) month++;
					while (month > 12) month--;
					while (day < 1) day++;
					while (day > 28) day--;
					while (minute < 0) minute++;
					while (minute > 59) minute--;
					while (second < 0) second++;
					while (second > 59) second--;
					d = new DateTime(year, month, day, hour, minute, second, 0);
					success = true;
				}
				catch (ArgumentOutOfRangeException) { }
			}
		}

		if (!success)
		{
			var msg = string.Format("y({0}) m({1}) d({2}) h({3}) m({4}) s({5})", year, month, day, hour, minute, second);
			throw new ZipException(string.Format("Bad date/time format in the zip file. ({0})", msg));

		}
		// workitem 6191
		//d = AdjustTime_Reverse(d);
		d = DateTime.SpecifyKind(d, DateTimeKind.Local);
		return d;
	}


	internal
	 static int DateTimeToPacked(DateTime time)
	{
		// The time is passed in here only for purposes of writing LastModified to the
		// zip archive. It should always be LocalTime, but we convert anyway.  And,
		// since the time is being written out, it needs to be adjusted.

		time = time.ToLocalTime();
		// workitem 7966
		//time = AdjustTime_Forward(time);

		// see http://www.vsft.com/hal/dostime.htm for the format
		var packedDate = (ushort)(time.Day & 0x0000001F | time.Month << 5 & 0x000001E0 | time.Year - 1980 << 9 & 0x0000FE00);
		var packedTime = (ushort)(time.Second / 2 & 0x0000001F | time.Minute << 5 & 0x000007E0 | time.Hour << 11 & 0x0000F800);

		var result = (int)((uint)(packedDate << 16) | packedTime);
		return result;
	}


	/// <summary>
	///   Create a pseudo-random filename, suitable for use as a temporary
	///   file, and open it.
	/// </summary>
	/// <remarks>
	/// <para>
	///   The System.IO.Path.GetRandomFileName() method is not available on
	///   the Compact Framework, so this library provides its own substitute
	///   on NETCF.
	/// </para>
	/// <para>
	///   This method produces a filename of the form
	///   DotNetZip-xxxxxxxx.tmp, where xxxxxxxx is replaced by randomly
	///   chosen characters, and creates that file.
	/// </para>
	/// </remarks>
	public static void CreateAndOpenUniqueTempFile(string dir,
												   out Stream fs,
												   out string filename)
	{
		// workitem 9763
		// http://dotnet.org.za/markn/archive/2006/04/15/51594.aspx
		// try 3 times:
		for (var i = 0; i < 3; i++)
		{
			try
			{
				filename = Path.Combine(dir, InternalGetTempFileName());
				fs = new FileStream(filename, FileMode.CreateNew);
				return;
			}
			catch (IOException)
			{
				if (i == 2) throw;
			}
		}

		throw new IOException();
	}


	public static string InternalGetTempFileName() => "DotNetZip-" + Path.GetRandomFileName()[..8] + ".tmp";


	/// <summary>
	/// Workitem 7889: handle ERROR_LOCK_VIOLATION during read
	/// </summary>
	/// <remarks>
	/// This could be gracefully handled with an extension attribute, but
	/// This assembly is built for .NET 2.0, so I cannot use them.
	/// </remarks>
	internal static int ReadWithRetry(Stream s, byte[] buffer, int offset, int count, string FileName)
	{
		var n = 0;
		var done = false;
		do
		{
			try
			{
				n = s.Read(buffer, offset, count);
				done = true;
			}
			catch /*(System.IO.IOException ioexc1)*/
			{
				// Check if we can call GetHRForException,
				// which makes unmanaged code calls.
				//var p = new SecurityPermission(SecurityPermissionFlag.UnmanagedCode);
				//if (p.IsUnrestricted())
				//{
				//    uint hresult = _HRForException(ioexc1);
				//    if (hresult != 0x80070021)  // ERROR_LOCK_VIOLATION
				//        throw new System.IO.IOException(String.Format("Cannot read file {0}", FileName), ioexc1);
				//    retries++;
				//    if (retries > 10)
				//        throw new System.IO.IOException(String.Format("Cannot read file {0}, at offset 0x{1:X8} after 10 retries", FileName, offset), ioexc1);

				//    // max time waited on last retry = 250 + 10*550 = 5.75s
				//    // aggregate time waited after 10 retries: 250 + 55*550 = 30.5s
				//    System.Threading.Thread.Sleep(250 + retries * 550);
				//}
				//else
				//{
				// The permission.Demand() failed. Therefore, we cannot call
				// GetHRForException, and cannot do the subtle handling of
				// ERROR_LOCK_VIOLATION.  Just bail.
				throw;
				//}
			}
		}
		while (!done);

		return n;
	}


	// workitem 8009
	//
	// This method must remain separate.
	//
	// Marshal.GetHRForException() is needed to do special exception handling for
	// the read.  But, that method requires UnmanagedCode permissions, and is marked
	// with LinkDemand for UnmanagedCode.  In an ASP.NET medium trust environment,
	// where UnmanagedCode is restricted, will generate a SecurityException at the
	// time of JIT of the method that calls a method that is marked with LinkDemand
	// for UnmanagedCode. The SecurityException, if it is restricted, will occur
	// when this method is JITed.
	//
	// The Marshal.GetHRForException() is factored out of ReadWithRetry in order to
	// avoid the SecurityException at JIT compile time. Because _HRForException is
	// called only when the UnmanagedCode is allowed.  This means .NET never
	// JIT-compiles this method when UnmanagedCode is disallowed, and thus never
	// generates the JIT-compile time exception.
	//
	private static uint _HRForException(Exception ex1) => unchecked((uint)System.Runtime.InteropServices.Marshal.GetHRForException(ex1));

}



/// <summary>
///   A decorator stream. It wraps another stream, and performs bookkeeping
///   to keep track of the stream Position.
/// </summary>
/// <remarks>
///   <para>
///     In some cases, it is not possible to get the Position of a stream, let's
///     say, on a write-only output stream like ASP.NET's
///     <c>Response.OutputStream</c>, or on a different write-only stream
///     provided as the destination for the zip by the application.  In this
///     case, programmers can use this counting stream to count the bytes read
///     or written.
///   </para>
///   <para>
///     Consider the scenario of an application that saves a self-extracting
///     archive (SFX), that uses a custom SFX stub.
///   </para>
///   <para>
///     Saving to a filesystem file, the application would open the
///     filesystem file (getting a <c>FileStream</c>), save the custom sfx stub
///     into it, and then call <c>ZipFile.Save()</c>, specifying the same
///     FileStream. <c>ZipFile.Save()</c> does the right thing for the zipentry
///     offsets, by inquiring the Position of the <c>FileStream</c> before writing
///     any data, and then adding that initial offset into any ZipEntry
///     offsets in the zip directory. Everything works fine.
///   </para>
///   <para>
///     Now suppose the application is an ASPNET application and it saves
///     directly to <c>Response.OutputStream</c>. It's not possible for DotNetZip to
///     inquire the <c>Position</c>, so the offsets for the SFX will be wrong.
///   </para>
///   <para>
///     The workaround is for the application to use this class to wrap
///     <c>HttpResponse.OutputStream</c>, then write the SFX stub and the ZipFile
///     into that wrapper stream. Because <c>ZipFile.Save()</c> can inquire the
///     <c>Position</c>, it will then do the right thing with the offsets.
///   </para>
/// </remarks>
internal class CountingStream : Stream
{
	// workitem 12374: this class is now public
	private readonly Stream _s;
	private long _bytesWritten;
	private long _bytesRead;
	private readonly long _initialOffset;

	/// <summary>
	/// The constructor.
	/// </summary>
	/// <param name="stream">The underlying stream</param>
	public CountingStream(Stream stream)
		: base()
	{
		_s = stream;
		try
		{
			_initialOffset = _s.Position;
		}
		catch
		{
			_initialOffset = 0L;
		}
	}

	/// <summary>
	///   Gets the wrapped stream.
	/// </summary>
	public Stream WrappedStream => _s;

	/// <summary>
	///   The count of bytes written out to the stream.
	/// </summary>
	public long BytesWritten => _bytesWritten;

	/// <summary>
	///   the count of bytes that have been read from the stream.
	/// </summary>
	public long BytesRead => _bytesRead;

	/// <summary>
	///    Adjust the byte count on the stream.
	/// </summary>
	///
	/// <param name='delta'>
	///   the number of bytes to subtract from the count.
	/// </param>
	///
	/// <remarks>
	///   <para>
	///     Subtract delta from the count of bytes written to the stream.
	///     This is necessary when seeking back, and writing additional data,
	///     as happens in some cases when saving Zip files.
	///   </para>
	/// </remarks>
	public void Adjust(long delta)
	{
		_bytesWritten -= delta;
		if (_bytesWritten < 0)
			throw new InvalidOperationException();
		if (_s as CountingStream != null)
			((CountingStream)_s).Adjust(delta);
	}

	/// <summary>
	///   The read method.
	/// </summary>
	/// <param name="buffer">The buffer to hold the data read from the stream.</param>
	/// <param name="offset">the offset within the buffer to copy the first byte read.</param>
	/// <param name="count">the number of bytes to read.</param>
	/// <returns>the number of bytes read, after decryption and decompression.</returns>
	public override int Read(byte[] buffer, int offset, int count)
	{
		var n = _s.Read(buffer, offset, count);
		_bytesRead += n;
		return n;
	}

	/// <summary>
	///   Write data into the stream.
	/// </summary>
	/// <param name="buffer">The buffer holding data to write to the stream.</param>
	/// <param name="offset">the offset within that data array to find the first byte to write.</param>
	/// <param name="count">the number of bytes to write.</param>
	public override void Write(byte[] buffer, int offset, int count)
	{
		if (count == 0) return;
		_s.Write(buffer, offset, count);
		_bytesWritten += count;
	}

	/// <summary>
	///   Whether the stream can be read.
	/// </summary>
	public override bool CanRead => _s.CanRead;

	/// <summary>
	///   Whether it is possible to call Seek() on the stream.
	/// </summary>
	public override bool CanSeek => _s.CanSeek;

	/// <summary>
	///   Whether it is possible to call Write() on the stream.
	/// </summary>
	public override bool CanWrite => _s.CanWrite;

	/// <summary>
	///   Flushes the underlying stream.
	/// </summary>
	public override void Flush() => _s.Flush();

	/// <summary>
	///   The length of the underlying stream.
	/// </summary>
	public override long Length => _s.Length; // BytesWritten?

	/// <summary>
	///   Returns the sum of number of bytes written, plus the initial
	///   offset before writing.
	/// </summary>
	public long ComputedPosition => _initialOffset + _bytesWritten;


	/// <summary>
	///   The Position of the stream.
	/// </summary>
	public override long Position
	{
		get { return _s.Position; }
		set
		{
			_s.Seek(value, SeekOrigin.Begin);
			// workitem 10178
			SharedUtilities.Workaround_Ladybug318918(_s);
		}
	}

	/// <summary>
	///   Seek in the stream.
	/// </summary>
	/// <param name="offset">the offset point to seek to</param>
	/// <param name="origin">the reference point from which to seek</param>
	/// <returns>The new position</returns>
	public override long Seek(long offset, SeekOrigin origin) => _s.Seek(offset, origin);

	/// <summary>
	///   Set the length of the underlying stream.  Be careful with this!
	/// </summary>
	///
	/// <param name='value'>the length to set on the underlying stream.</param>
	public override void SetLength(long value) => _s.SetLength(value);
}