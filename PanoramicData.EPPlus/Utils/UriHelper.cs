using System;
using System.Text;

namespace OfficeOpenXml.Utils;

internal static class UriHelper
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

		for (var i = s; i >= 0; i--)
		{
			file = source[i] + "/" + file;
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

		var relative = new StringBuilder();
		for (var s = i; s < slen; s++)
		{
			relative.Append("../");
		}

		var fileStart = relative.Length;
		for (var t = i; t < target.Length; t++)
		{
			if (relative.Length > fileStart)
			{
				relative.Append('/');
			}

			relative.Append(target[t]);
		}

		return new Uri(relative.ToString(), UriKind.Relative);
	}
}
