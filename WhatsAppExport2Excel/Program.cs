using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace WhatsAppChatToExcel
{
    internal sealed class ChatMessage
    {
        public DateTime Date { get; set; }
        public string Sender { get; set; } = "";
        public string Message { get; set; } = "";
        public string Media { get; set; } = "";
        public bool IsSystem { get; set; }
    }

    internal static class Program
    {
        // === CLI options ===
        private static string? _mediaDir;
        private static bool _skipSystem;
        private static string? _cultureName;
        private static TimeSpan? _tzOffset;
        private static bool _forceRtl;
        private static DateTime? _fromDate;
        private static DateTime? _toDate;

        // progress tracking
        private static int _lastShownPercent = -1;

        // Regex patterns for chat lines
        private static readonly Regex[] LineStartPatterns =
        {
            new Regex(@"^(?<d>\d{1,2}/\d{1,2}/\d{2,4}),\s*(?<t>\d{1,2}:\d{2}(?::\d{2})?)\s*(?<ampm>(?:[AaPp]\.?[Mm]\.?)?)\s*[-–]\s*(?<name>.+?):\s*(?<msg>.*)$",
                RegexOptions.Compiled),
            new Regex(@"^\[\s*(?<d>\d{1,2}/\d{1,2}/\d{2,4}),\s*(?<t>\d{1,2}:\d{2}(?::\d{2})?)\s*(?<ampm>(?:[AaPp]\.?[Mm]\.?)?)\s*\]\s*(?<name>.+?):\s*(?<msg>.*)$",
                RegexOptions.Compiled),
        };

        private static readonly string[] TimestampFormats =
        {
            "d/M/yy, H:mm", "d/M/yy, H:mm:ss",
            "d/M/yyyy, H:mm", "d/M/yyyy, H:mm:ss",
            "d/M/yy, h:mm tt", "d/M/yy, h:mm:ss tt",
            "d/M/yyyy, h:mm tt", "d/M/yyyy, h:mm:ss tt",
            "M/d/yy, H:mm", "M/d/yy, H:mm:ss",
            "M/d/yyyy, H:mm", "M/d/yyyy, H:mm:ss",
            "M/d/yy, h:mm tt", "M/d/yy, h:mm:ss tt",
            "M/d/yyyy, h:mm tt", "M/d/yyyy, h:mm:ss tt",
        };

        private static readonly Regex FileLike =
            new Regex(@"\b[\p{L}\p{Nd}_\-]+\.(jpg|jpeg|png|gif|mp4|mp3|opus|pdf|docx?|xlsx?|pptx?|heic|mov|zip)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly string[] SystemStarts =
        {
            "messages to this chat are now",
            "security code changed",
            "missed voice call", "missed video call",
            "تم إنشاء المجموعة", "قام بتغيير صورة المجموعة", "قام بتغيير وصف المجموعة",
            "تم تغيير رقم الهاتف", "أصبحت الرسائل الآن"
        };

        // === Helpers ===
        private static bool ContainsArabic(string s) =>
            s.AsSpan().IndexOfAnyInRange('\u0600', '\u06FF') >= 0
            || s.AsSpan().IndexOfAnyInRange('\u0750', '\u077F') >= 0
            || s.AsSpan().IndexOfAnyInRange('\u08A0', '\u08FF') >= 0;

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
            return SystemStarts.Any(p => t.StartsWith(p, StringComparison.OrdinalIgnoreCase));
        }

        private static string DetectMedia(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            if (text.Contains("<Media omitted>", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("المرفق غير متاح", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("image omitted", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("(file attached)", StringComparison.OrdinalIgnoreCase))
            {
                var m = FileLike.Match(text);
                return m.Success ? m.Value : "Yes";
            }
            var match = FileLike.Match(text);
            return match.Success ? match.Value : "";
        }

        private static bool TryParseHeader(string rawLine, CultureInfo culture, out DateTime timestamp, out string sender, out string firstMessage)
        {
            timestamp = default;
            sender = "";
            firstMessage = "";
            var line = NormalizeDigitsAndSpaces(rawLine);

            foreach (var rx in LineStartPatterns)
            {
                var m = rx.Match(line);
                if (!m.Success) continue;

                var d = m.Groups["d"].Value.Trim();
                var t = m.Groups["t"].Value.Trim();
                var ampmRaw = m.Groups["ampm"].Value.Trim();
                string? ampm = string.IsNullOrEmpty(ampmRaw) ? null : ampmRaw.Replace(".", "").ToUpperInvariant();
                string stamp = ampm == null ? $"{d}, {t}" : $"{d}, {t} {ampm}";

                if (DateTime.TryParseExact(stamp, TimestampFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out timestamp) ||
                    DateTime.TryParse(stamp, culture, DateTimeStyles.None, out timestamp))
                {
                    sender = m.Groups["name"].Value.Trim();
                    firstMessage = m.Groups["msg"].Value;
                    return true;
                }
            }
            return false;
        }

        private static void ShowProgress(long pos, long len)
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

        private static IEnumerable<ChatMessage> ParseChat(string path, CultureInfo culture)
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            long total = fs.Length;
            using var reader = new StreamReader(fs, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);

            DateTime currentTs = default;
            string currentSender = "";
            var sb = new StringBuilder();
            string currentMedia = "";
            bool hasCurrent = false;
            bool currentSystem = false;

            string? line;
            while ((line = reader.ReadLine()) is not null)
            {
                ShowProgress(reader.BaseStream.Position, total);

                if (TryParseHeader(line, culture, out var ts, out var sender, out var first))
                {
                    if (hasCurrent)
                        yield return new ChatMessage { Date = currentTs, Sender = currentSender, Message = sb.ToString().TrimEnd(), Media = currentMedia, IsSystem = currentSystem };

                    currentTs = ts;
                    currentSender = sender;
                    sb.Clear();
                    sb.Append(first);
                    currentMedia = DetectMedia(first);
                    currentSystem = LooksSystemMessage(first);
                    hasCurrent = true;
                }
                else if (hasCurrent)
                {
                    sb.AppendLine();
                    sb.Append(line);
                    var maybe = DetectMedia(line);
                    if (!string.IsNullOrEmpty(maybe) && string.IsNullOrEmpty(currentMedia))
                        currentMedia = maybe;
                    if (!currentSystem && LooksSystemMessage(line))
                        currentSystem = true;
                }
            }

            if (hasCurrent)
                yield return new ChatMessage { Date = currentTs, Sender = currentSender, Message = sb.ToString().TrimEnd(), Media = currentMedia, IsSystem = currentSystem };

            ShowProgress(total, total);
        }

        private static DateTimeOffset ApplyTimezone(DateTime dt, TimeSpan? tzOffset) =>
            tzOffset == null ? new DateTimeOffset(dt) : new DateTimeOffset(DateTime.SpecifyKind(dt, DateTimeKind.Unspecified), tzOffset.Value);

        private static string? ResolveMediaLink(string dir, string token)
        {
            var exact = Path.Combine(dir, token);
            if (File.Exists(exact)) return exact;
            var nameNoExt = Path.GetFileNameWithoutExtension(token);
            var ext = Path.GetExtension(token);
            return Directory.EnumerateFiles(dir, $"*{nameNoExt}*{ext}", SearchOption.AllDirectories).FirstOrDefault();
        }

        private static void WriteExcelStreaming(string inputPath, string outputPath, CultureInfo culture)
        {
            using var wb = new XLWorkbook();
            var sheets = new Dictionary<DateTime, (IXLWorksheet ws, int nextRow, int arabicScore)>();

            void EnsureSheet(DateTime day)
            {
                if (sheets.ContainsKey(day)) return;
                var ws = wb.Worksheets.Add(day.ToString("yyyy-MM-dd"));
                ws.Cell(1, 1).Value = "Date";
                ws.Cell(1, 2).Value = "Sender";
                ws.Cell(1, 3).Value = "Message";
                ws.Cell(1, 4).Value = "Media";
                ws.Range(1, 1, 1, 4).Style.Font.Bold = true;
                ws.SheetView.FreezeRows(1);
                ws.Column(1).Width = 20;
                ws.Column(2).Width = 28;
                ws.Column(3).Width = 90;
                ws.Column(4).Width = 30;
                sheets[day] = (ws, 2, 0);
            }

            foreach (var m in ParseChat(inputPath, culture))
            {
                if (_skipSystem && m.IsSystem) continue;
                var dto = ApplyTimezone(m.Date, _tzOffset);
                var day = dto.Date;

                if (_fromDate.HasValue && day < _fromDate.Value) continue;
                if (_toDate.HasValue && day > _toDate.Value) continue;

                EnsureSheet(day);
                var entry = sheets[day];
                var ws = entry.ws;
                int row = entry.nextRow;

                ws.Cell(row, 1).Value = dto.LocalDateTime;
                ws.Cell(row, 1).Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
                ws.Cell(row, 2).Value = m.Sender;
                ws.Cell(row, 3).Value = m.Message;
                ws.Cell(row, 3).Style.Alignment.WrapText = true;

                if (!string.IsNullOrEmpty(m.Media) && !string.IsNullOrEmpty(_mediaDir))
                {
                    var link = ResolveMediaLink(_mediaDir!, m.Media);
                    if (link != null)
                    {
                        ws.Cell(row, 4).Value = Path.GetFileName(link);
                        var hl = ws.Cell(row, 4).CreateHyperlink();
                        hl.ExternalAddress = new Uri(link, UriKind.Absolute);
                    }
                    else ws.Cell(row, 4).Value = m.Media;
                }
                else ws.Cell(row, 4).Value = m.Media;

                int score = entry.arabicScore;
                if (ContainsArabic(m.Message) || ContainsArabic(m.Sender)) score++;
                sheets[day] = (ws, row + 1, score);
            }

            foreach (var kv in sheets)
                if (_forceRtl || kv.Value.arabicScore >= 20)
                    kv.Value.ws.RightToLeft = true;

            wb.SaveAs(outputPath);
        }

        private static void ParseArgs(string[] args, out string input, out string output)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage:\n  WhatsAppChatToExcel <input_chat.txt> <output.xlsx> [--media-dir=...] [--skip-system] [--culture=ar-SA] [--timezone=+03:00] [--rtl] [--from=YYYY-MM-DD] [--to=YYYY-MM-DD]");
                Environment.Exit(1);
            }
            input = Path.GetFullPath(args[0]);
            output = Path.GetFullPath(args[1]);
            foreach (var a in args.Skip(2))
            {
                if (a.StartsWith("--media-dir=")) _mediaDir = a.Split('=', 2)[1].Trim('"');
                else if (a == "--skip-system") _skipSystem = true;
                else if (a.StartsWith("--culture=")) _cultureName = a.Split('=', 2)[1];
                else if (a.StartsWith("--timezone=") && TimeSpan.TryParse(a.Split('=', 2)[1], out var tz)) _tzOffset = tz;
                else if (a == "--rtl") _forceRtl = true;
                else if (a.StartsWith("--from=") && DateTime.TryParseExact(a.Split('=', 2)[1], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d1)) _fromDate = d1.Date;
                else if (a.StartsWith("--to=") && DateTime.TryParseExact(a.Split('=', 2)[1], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d2)) _toDate = d2.Date;
            }
        }

        private static void Main(string[] args)
        {
            ParseArgs(args, out var input, out var output);
            var culture = CultureInfo.InvariantCulture;
            if (!string.IsNullOrWhiteSpace(_cultureName))
                try { culture = CultureInfo.GetCultureInfo(_cultureName); } catch { }
            WriteExcelStreaming(input, output, culture);
            Console.WriteLine($"\nExcel written: {output}");
            if (_fromDate.HasValue || _toDate.HasValue)
                Console.WriteLine($"Filtered days: {(_fromDate?.ToString("yyyy-MM-dd") ?? "min")} → {(_toDate?.ToString("yyyy-MM-dd") ?? "max")}");
        }
    }

    internal static class SpanExtensions
    {
        public static int IndexOfAnyInRange(this ReadOnlySpan<char> span, char minInclusive, char maxInclusive)
        {
            for (int i = 0; i < span.Length; i++)
                if (span[i] >= minInclusive && span[i] <= maxInclusive) return i;
            return -1;
        }
    }
}
