using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace WhatsAppChatToExcel;
internal sealed class ChatParser
{
    private readonly ChatParserOptions _options;
    private int _lastShownPercent = -1;
    private readonly HashSet<string> _usedMediaFiles = new();

    public ChatParser(ChatParserOptions options)
    {
        _options = options;
    }

    public IEnumerable<ChatMessage> ParseChat()
    {
        var path = _options.InputPath;
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        long total = fs.Length;
        using var reader = new StreamReader(fs, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);

        DateTime currentTs = default;
        string currentSender = "";
        var sb = new StringBuilder();
        bool hasCurrent = false;
        bool currentSystem = false;

        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            ShowProgress(reader.BaseStream.Position, total);

            if (TryParseHeader(line, out var ts, out var sender, out var first))
            {
                if (hasCurrent)
                {
                    var message = sb.ToString().TrimEnd();
                    var chatMessage = new ChatMessage { Date = currentTs, Sender = currentSender, Message = message, IsSystem = currentSystem };
                    if (IsMatched(chatMessage))
                    {
                        TrackUsedMediaFiles(message);
                        yield return chatMessage;
                    }
                }

                currentTs = ts;
                currentSender = sender;
                sb.Clear();
                sb.Append(first);
                currentSystem = LooksSystemMessage(first);
                hasCurrent = true;
            }
            else if (hasCurrent)
            {
                sb.AppendLine();
                sb.Append(line);
                if (!currentSystem && LooksSystemMessage(line))
                    currentSystem = true;
            }
        }

        if (hasCurrent)
        {
            var message = sb.ToString().TrimEnd();
            
            var chatMessage = new ChatMessage { Date = currentTs, Sender = currentSender, Message = message, IsSystem = currentSystem };
            if (IsMatched(chatMessage))
            {
                TrackUsedMediaFiles(message);
                yield return chatMessage;
            }
        }

        ShowProgress(total, total);

        if (_options.MoveNotUsedMedia)
        {
            MoveUnusedMediaFiles();
        }
    }

    private bool IsMatched(ChatMessage message)
    {
        if (_options.SkipSystem && message.IsSystem) return false;

        var dto = new DateTimeOffset(message.Date);
        var day = dto.Date;

        if (_options.FromDate.HasValue && day < _options.FromDate.Value) return false;
        if (_options.ToDate.HasValue && day > _options.ToDate.Value) return false;

        return true;
    }

    private void TrackUsedMediaFiles(string message)
    {
        var token = MediaHelper.DetectMediaToken(message);
        if (!string.IsNullOrEmpty(token))
        {
            var link = MediaHelper.ResolveMediaLink(_options.MediaDirectory!, token);
            _usedMediaFiles.Add(Path.GetFileName(link)!);
        }
    }

    private void MoveUnusedMediaFiles()
    {
        if (string.IsNullOrEmpty(_options.MediaDirectory) || !Directory.Exists(_options.MediaDirectory))
            return;

        var mediaFiles =
            Directory.GetFiles(_options.MediaDirectory, "*.*", SearchOption.TopDirectoryOnly)
            .Select(Path.GetFileName)
            .ToArray();

        if (mediaFiles.Length == 0)
            return;

        var notUsedFolder = Path.Join(_options.MediaDirectory, "NotUsed");

        var unusedFiles = mediaFiles.Except(_usedMediaFiles).ToArray();

        if (unusedFiles.Length == 0)
            return;

        Directory.CreateDirectory(notUsedFolder);

        foreach (var file in unusedFiles)
        {
            try
            {
                var sourcePath = Path.Combine(_options.MediaDirectory, file);
                var destPath = Path.Combine(notUsedFolder, file);
                File.Move(sourcePath, destPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not move file {file}: {ex.Message}");
            }
        }

        Console.WriteLine($"Moved {unusedFiles.Length} unused media files to NotUsed folder.");
    }

    private bool TryParseHeader(string rawLine, out DateTime timestamp, out string sender, out string firstMessage)
    {
        timestamp = default;
        sender = "";
        firstMessage = "";
        var line = NormalizeDigitsAndSpaces(rawLine);

        foreach (var rx in ChatParserOptions.LineStartPatterns)
        {
            var m = rx.Match(line);
            if (!m.Success) continue;

            var d = m.Groups["d"].Value.Trim();
            var t = m.Groups["t"].Value.Trim();
            var ampmRaw = m.Groups["ampm"].Value.Trim();
            string? ampm = string.IsNullOrEmpty(ampmRaw) ? null : ampmRaw.Replace(".", "").ToUpperInvariant();
            string stamp = ampm == null ? $"{d}, {t}" : $"{d}, {t} {ampm}";

            if (DateTime.TryParseExact(stamp, ChatParserOptions.TimestampFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out timestamp) ||
                DateTime.TryParse(stamp, CultureInfo.InvariantCulture, DateTimeStyles.None, out timestamp))
            {
                sender = m.Groups["name"].Value.Trim();
                firstMessage = m.Groups["msg"].Value;
                return true;
            }
        }
        return false;
    }

    private static string NormalizeDigitsAndSpaces(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s)
        {
            if (ch >= '\u0660' && ch <= '\u0669') { sb.Append((char)('0' + (ch - '\u0660'))); continue; }
            if (ch >= '\u06F0' && ch <= '\u06F9') { sb.Append((char)('0' + (ch - '\u06F0'))); continue; }
            if (ch == '\u00A0' || ch == '\u202F' || ch == '\u2007' || ch == '\u2060') { sb.Append(' '); continue; }
            sb.Append(ch);
        }
        return Regex.Replace(sb.ToString(), @"\s+(?=[AaPp]\.?[Mm]\.?)", " ");
    }

    private static bool LooksSystemMessage(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var t = text.Trim();
        return ChatParserOptions.SystemMessageStarts.Any(p => t.StartsWith(p, StringComparison.OrdinalIgnoreCase));
    }

    private void ShowProgress(long pos, long len)
    {
        if (len <= 0) return;
        int pct = (int)(pos * 100 / len);
        if (pct != _lastShownPercent)
        {
            _lastShownPercent = pct;
            Console.Write($"\rParsing... {pct}%");
            if (pct == 100) Console.WriteLine();
        }
    }
}
