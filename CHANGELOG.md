# Changelog

## 11.0

11.0 is a major release with breaking changes.

### Breaking changes

- **Targets .NET 11 (`net11.0`).** The previous release, 10.1, targeted .NET 10.
- **System.Drawing.Common (GDI+) is no longer used**, so the library runs the same on Windows, Linux and macOS.
  - Removed: `ExcelDrawings.AddPicture(string, Image[, Uri])`; `ExcelPicture.Image` and `ExcelPicture.ImageFormat`; `ExcelBackgroundImage.Image`; `ExcelVmlDrawingPicture.Image`; `ExcelHeaderFooterText.InsertPicture(Image, PictureAlignment)`; `ExcelFont.SetFromFont(Font)`, `ExcelTextFont.SetFromFont(Font)` and `ExcelFontXml.SetFromFont(Font)`; the `[SupportedOSPlatform("windows")]` attributes.
  - Added: `ExcelDrawings.AddPicture(string, byte[])` and `AddPicture(string, Stream)` (with optional hyperlink); `ExcelPicture.ImageBytes` and `ExcelPicture.ContentType`; `ExcelBackgroundImage.ImageBytes` and `SetFromStream`; `ExcelVmlDrawingPicture.ImageBytes`; `ExcelHeaderFooterText.InsertPicture(byte[], ...)` and `InsertPicture(Stream, ...)`. Image size and resolution are read from the PNG, JPEG, GIF, BMP, TIFF, EMF or WMF header and the file bytes are stored unchanged (pictures are no longer re-encoded as JPEG).
  - `AutoFitColumns` measures text with SkiaSharp (MIT) instead of GDI+. Results can differ slightly from earlier versions, and on a machine without the named font the platform default font is measured.
- **The bundled DotNetZip code has been removed.** Packages are read and written with `System.IO.Compression` (`ZipArchive`). The following public types are gone; use `System.IO.Compression` instead (`DeflateStream`, `GZipStream`, `ZLibStream`, `ZipArchive`, `CompressionLevel`, `CompressionMode`, `InvalidDataException`):
  - `OfficeOpenXml.Packaging.DotNetZip`: `WriteDelegate`, `OpenDelegate`, `CloseDelegate`, `SetCompressionCallback`, `ZipException`, `BadPasswordException`, `BadReadException`, `BadCrcException`, `SfxGenerationException`, `BadStateException`.
  - `OfficeOpenXml.Packaging.DotNetZip.Zlib`: `DeflateStream`, `GZipStream`, `ZlibStream`, `ParallelDeflateOutputStream`, `FlushType`, `CompressionLevel`, `CompressionStrategy`, `CompressionMode`, `ZlibException`, `Adler`, `ZlibCodec`, `ZlibConstants`.
- **Specific exception types.** Some methods in Drawing, Style and ConditionalFormatting now throw `InvalidOperationException` or `ArgumentException` instead of `System.Exception`, with the same messages. Code that catches `Exception` is unaffected.

### Behaviour notes

- `ExcelPackage.Compression` (`OfficeOpenXml.CompressionLevel`) is unchanged as an API. Its values now map to `System.IO.Compression.CompressionLevel`: `Level1` to `Level3` to `Fastest`, `Level9` to `SmallestSize`, and the others to `Optimal`. As before, `None` still deflates entries rather than storing them. Compressed sizes differ slightly from 10.x because the deflate implementation is different; part names, part order and content types are unchanged.
- Zip64 is still used automatically when, and only when, a workbook exceeds the classic ZIP limits (4 GB, or 65,534 entries).
- Opening a stream that is not a zip package still throws `InvalidDataException` with the same message; the underlying `System.IO.Compression` error is now attached as the inner exception.
- On .NET 11 the runtime converts `double` and `float` to `decimal` exactly instead of rounding to 15 (double) or 7 (float) significant digits. EPPlus keeps the earlier rounding, so `GetValue<decimal>()` on a cell holding `0.396180555555556` still returns `0.396180555555556`.

## 10.1

- Large workbooks can be saved. A workbook that exceeded the classic ZIP limits (4 GB of size or offset, or 65,534 internal entries) previously threw while saving. Zip64 is now used automatically when, and only when, a workbook needs it.
- This release targets .NET 10. The previous release, 10.1.1, targeted .NET 9.
