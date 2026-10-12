using SkiaSharp;
using System.Drawing;

namespace OfficeOpenXml.Compatibility;

/// <summary>
/// Measures text for column auto-fit, using SkiaSharp so that it works on Windows, Linux and macOS.
/// </summary>
internal sealed class TextMeasurer : IDisposable
{
	private const float PixelsPerPoint = 96F / 72F;

	private static readonly Lazy<bool> _isAvailable = new(Probe);

	private readonly Dictionary<(string Name, bool Bold, bool Italic), SKTypeface> _typefaces = [];
	private readonly Dictionary<(string Name, float Size, bool Bold, bool Italic), SKFont> _fonts = [];

	/// <summary>
	/// False when the native SkiaSharp library cannot be loaded on this machine.
	/// </summary>
	internal static bool IsAvailable => _isAvailable.Value;

	private static bool Probe()
	{
		try
		{
			using var typeface = SKTypeface.Default;
			using var font = new SKFont(typeface, 12);
			font.MeasureText("0");
			return true;
		}
		catch (Exception ex) when (ex is DllNotFoundException or TypeInitializationException or EntryPointNotFoundException or BadImageFormatException)
		{
			return false;
		}
	}

	/// <summary>
	/// The size in pixels (at 96 DPI) of the text in the given font. Size is in points.
	/// A font that is not installed falls back to the platform's default typeface.
	/// </summary>
	internal SizeF Measure(string text, string fontName, float sizeInPoints, bool bold, bool italic)
	{
		if (string.IsNullOrEmpty(text))
		{
			return SizeF.Empty;
		}

		var key = (fontName, sizeInPoints, bold, italic);
		if (!_fonts.TryGetValue(key, out var font))
		{
			if (!_typefaces.TryGetValue((fontName, bold, italic), out var typeface))
			{
				typeface = SKTypeface.FromFamilyName(
					fontName,
					bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
					SKFontStyleWidth.Normal,
					italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright)
					?? SKTypeface.Default;
				_typefaces.Add((fontName, bold, italic), typeface);
			}

			font = new SKFont(typeface, sizeInPoints * PixelsPerPoint);
			_fonts.Add(key, font);
		}

		var width = font.MeasureText(text);
		var height = font.Spacing;
		return new SizeF(width, height);
	}

	public void Dispose()
	{
		foreach (var font in _fonts.Values)
		{
			font.Dispose();
		}

		foreach (var typeface in _typefaces.Values)
		{
			typeface.Dispose();
		}

		_fonts.Clear();
		_typefaces.Clear();
	}
}
