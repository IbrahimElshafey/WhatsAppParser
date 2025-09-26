using System;
using System.Linq;

namespace WhatsAppChatToExcel;

internal static class MediaHelper
{
    public static string DetectMediaToken(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        if (text.Contains("<Media omitted>", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("المرفق غير متاح", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("image omitted", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("(file attached)", StringComparison.OrdinalIgnoreCase))
        {
            var m = ChatParserOptions.FileLikePattern.Match(text);
            return m.Success ? m.Value : "";
        }
        var match = ChatParserOptions.FileLikePattern.Match(text);
        return match.Success ? match.Value : "";
    }

    public static string? ResolveMediaLink(string dir, string token)
    {
        var exact = Path.Combine(dir, token);
        if (File.Exists(exact)) return exact;
        var nameNoExt = Path.GetFileNameWithoutExtension(token);
        var ext = Path.GetExtension(token);
        try
        {
            return Directory.EnumerateFiles(dir, $"*{nameNoExt}*{ext}", SearchOption.AllDirectories)
                            .OrderByDescending(File.GetCreationTimeUtc)
                            .FirstOrDefault();
        }
        catch { return null; }
    }
}
