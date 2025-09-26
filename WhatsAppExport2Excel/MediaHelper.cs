using System;
using System.Linq;

namespace WhatsAppChatToExcel;

internal static class MediaHelper
{
    public static string DetectMediaToken(string text)
    {
        // Return empty string if input text is null or whitespace
        if (string.IsNullOrWhiteSpace(text)) return "";
        
        // Check for known media placeholder texts in multiple languages
        if (text.Contains("<Media omitted>", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("المرفق غير متاح", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("image omitted", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("(file attached)", StringComparison.OrdinalIgnoreCase))
        {
            // Extract the actual filename from the media placeholder text
            var mediaPlaceholderMatch = ChatParserOptions.FileLikePattern.Match(text);
            return mediaPlaceholderMatch.Success ? mediaPlaceholderMatch.Value : "";
        }
        
        // For regular text, attempt to find any filename pattern
        var filenameMatch = ChatParserOptions.FileLikePattern.Match(text);
        return filenameMatch.Success ? filenameMatch.Value : "";
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
