﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Logging;

internal class TextFileLogger : IFormulaParserLogger
{
	private readonly StreamWriter _sw;
	private const string Separator = "=================================";
	private int _count;
	private readonly DateTime _startTime = DateTime.Now;
	private readonly Dictionary<string, int> _funcs = [];
	private readonly Dictionary<string, long> _funcPerformance = [];
	internal TextFileLogger(FileInfo fileInfo)
	{
		_sw = new StreamWriter(new FileStream(fileInfo.FullName, FileMode.Append));
	}

	private void WriteSeparatorAndTimeStamp()
	{
		_sw.WriteLine(Separator);
		_sw.WriteLine("Timestamp: {0}", DateTime.Now);
		_sw.WriteLine();
	}

	private void WriteAddressInfo(ParsingContext context)
	{
		if (context.Scopes.Current != null && context.Scopes.Current.Address != null)
		{
			_sw.WriteLine("Worksheet: {0}", context.Scopes.Current.Address.Worksheet ?? "<not specified>");
			_sw.WriteLine("Address: {0}", context.Scopes.Current.Address.Address ?? "<not available>");
		}
	}

	public void Log(ParsingContext context, Exception ex)
	{
		WriteSeparatorAndTimeStamp();
		WriteAddressInfo(context);
		_sw.WriteLine(ex);
		_sw.WriteLine();
	}

	public void Log(ParsingContext context, string message)
	{
		WriteSeparatorAndTimeStamp();
		WriteAddressInfo(context);
		_sw.WriteLine(message);
		_sw.WriteLine();
	}

	public void Log(string message)
	{
		WriteSeparatorAndTimeStamp();
		_sw.WriteLine(message);
		_sw.WriteLine();
	}

	public void LogCellCounted()
	{
		_count++;
		if (_count % 500 == 0)
		{
			_sw.WriteLine(Separator);
			var timeEllapsed = DateTime.Now.Subtract(_startTime);
			_sw.WriteLine("{0} cells parsed, time {1} seconds", _count, timeEllapsed.TotalSeconds);

			var funcs = _funcs.Keys.OrderByDescending(x => _funcs[x]).ToList();
			foreach (var func in funcs)
			{
				_sw.Write(func + "  - " + _funcs[func]);
				if (_funcPerformance.TryGetValue(func, out var value))
				{
					_sw.Write(" - avg: " + value / _funcs[func] + " milliseconds");
				}

				_sw.WriteLine();
			}

			_sw.WriteLine();
			_funcs.Clear();

		}
	}

	public void LogFunction(string func)
	{
		if (!_funcs.TryGetValue(func, out var value))
		{
			value = 0;
			_funcs.Add(func, value);
		}

		_funcs[func] = ++value;
	}

	public void LogFunction(string func, long milliseconds)
	{
		if (!_funcPerformance.ContainsKey(func))
		{
			_funcPerformance[func] = 0;
		}

		_funcPerformance[func] += milliseconds;
	}

	public void Dispose()
	{
		_sw.Close();
		_sw.Dispose();
	}
}
