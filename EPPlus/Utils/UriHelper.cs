using System;

namespace OfficeOpenXml.Utils;

internal class UriHelper
{
	internal static Uri ResolvePartUri(Uri sourceUri, Uri targetUri)
	{
		if (targetUri.OriginalString.StartsWith("/") || targetUri.OriginalString.Contains("://"))
		{
			return targetUri;
		}

		var source = sourceUri.OriginalString.Split('/');
		var target = targetUri.OriginalString.Split('/');

		var t = target.Length - 1;
		int s;
		if (sourceUri.OriginalString.EndsWith("/")) //is the source a directory?
		{
			s = source.Length - 1;
		}
		else
		{
			s = source.Length - 2;
		}

		var file = target[t--];

		while (t >= 0)
		{
			if (target[t] == ".")
			{
				break;
			}
			else if (target[t] == "..")
			{
				s--;
				t--;
			}
			else
			{
				file = target[t--] + "/" + file;
			}
		}

		if (s >= 0)
		{
			for (var i = s; i >= 0; i--)
			{
				file = source[i] + "/" + file;
			}
		}

		return new Uri(file, UriKind.RelativeOrAbsolute);
	}

	internal static Uri GetRelativeUri(Uri WorksheetUri, Uri uri)
	{
		var source = WorksheetUri.OriginalString.Split('/');
		var target = uri.OriginalString.Split('/');

		int slen;
		if (WorksheetUri.OriginalString.EndsWith("/"))
		{
			slen = source.Length;
		}
		else
		{
			slen = source.Length - 1;
		}

		var i = 0;
		while (i < slen && i < target.Length && source[i] == target[i])
		{
			i++;
		}

		var dirUp = "";
		for (var s = i; s < slen; s++)
		{
			dirUp += "../";
		}

		var file = "";
		for (var t = i; t < target.Length; t++)
		{
			file += (file == "" ? "" : "/") + target[t];
		}

		return new Uri(dirUp + file, UriKind.Relative);
	}
}
