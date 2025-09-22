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
        public DateTime Date { get; set; }          // UTC-normalized or local after TZ adjust
        public string Sender { get; set; } = "";
        public string Message { get; set; } = "";
        public string Media { get; set; } = "";
        public bool IsSystem { get; set; }
    }

    internal static class Program
    {
        // === CLI flags (simple parser) ======================================
        private static string? _mediaDir;
        private static bool _skipSystem;
        private static string? _cultureName;
        private static TimeSpan? _tzOffset; // e.g., +03:00
        private static bool _forceRtl;

        // === Regex upgrades ==================================================
        // We normalize digits and whitespace first; then these patterns match.
        //
        // Examples matched:
        //  "19/07/2025, 9:46 am - Name: Message"
        //  "19/07/2025, 10:22 am - Name: ...."
        //  "20/07/2025, 10:50 pm - Name: multi-line..."
        //  "[19/07/2025, 9:46 am] Name: Message"
        //
        // Notes:
        // - Accept hyphen (-) or en-dash (–) as the separator.
        // - Accept optional seconds.
        // - Accept AM/PM with or without dots and any spaces before it.
        // - Sender can contain commas and Arabic letters.
        private static readonly Regex[] LineStartPatterns =
        {
            // Android-like: d/M/yy(yy), h:mm[:ss] [am|pm] - Name: Message
            new Regex(@"^(?<d>\d{1,2}/\d{1,2}/\d{2,4}),\s*(?<t>\d{1,2}:\d{2}(?::\d{2})?)\s*(?<ampm>(?:[AaPp]\.?[Mm]\.?)?)\s*[-–]\s*(?<name>.+?):\s*(?<msg>.*)$",
                RegexOptions.Compiled),

            // iOS-like: [d/M/yy(yy), h:mm[:ss] [AM|PM]] Name: Message
            new Regex(@"^\[\s*(?<d>\d{1,2}/\d{1,2}/\d{2,4}),\s*(?<t>\d{1,2}:\d{2}(?::\d{2})?)\s*(?<ampm>(?:[AaPp]\.?[Mm]\.?)?)\s*\]\s*(?<name>.+?):\s*(?<msg>.*)$",
                RegexOptions.Compiled),
        };

        // Try these explicit formats first (invariant), then fall back to CultureInfo if provided.
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

        // File-like token to populate Media column.
        private static readonly Regex FileLike =
            new Regex(@"\b[\p{L}\p{Nd}_\-]+\.(jpg|jpeg|png|gif|mp4|mp3|opus|pdf|docx?|xlsx?|pptx?|heic|mov|zip)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // System message heuristics (English + Arabic common ones)
        private static readonly string[] SystemStarts =
        {
            "messages to this chat are now", // E2E notice (en)
            "security code changed",         // en
            "missed voice call", "missed video call",
            "تم إنشاء المجموعة", "قام بتغيير صورة المجموعة", "قام بتغيير وصف المجموعة",
            "تم تغيير رقم الهاتف", "أصبحت الرسائل الآن"
        };

        // Arabic ranges for quick RTL detection
        private static bool ContainsArabic(string s) =>
            s.AsSpan().IndexOfAnyInRange('\u0600', '\u06FF') >= 0
            || s.AsSpan().IndexOfAnyInRange('\u0750', '\u077F') >= 0
            || s.AsSpan().IndexOfAnyInRange('\u08A0', '\u08FF') >= 0;

        // Normalize Arabic-Indic digits and exotic spaces (e.g., narrow no-break space U+202F)
        private static string NormalizeDigitsAndSpaces(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;

            var sb = new StringBuilder(s.Length);
            foreach (var ch in s)
            {
                // Arabic-Indic 0-9
                if (ch >= '\u0660' && ch <= '\u0669') { sb.Append((char)('0' + (ch - '\u0660'))); continue; }
                // Extended Arabic-Indic 0-9
                if (ch >= '\u06F0' && ch <= '\u06F9') { sb.Append((char)('0' + (ch - '\u06F0'))); continue; }
                // Common weird spaces -> plain space
                if (ch == '\u00A0' || ch == '\u202F' || ch == '\u2007' || ch == '\u2060') { sb.Append(' '); continue; }
                sb.Append(ch);
            }
            // collapse multiple spaces around AM/PM just in case
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

            // WhatsApp generic markers (en/ar)
            if (text.IndexOf("<Media omitted>", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("المرفق غير متاح", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("image omitted", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("(file attached)", StringComparison.OrdinalIgnoreCase) >= 0)
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
                var ampmRaw = m.Groups["ampm"].Success ? m.Groups["ampm"].Value.Trim() : null;

                // Normalize AM/PM tokens (remove dots, uppercase)
                string? ampm = ampmRaw is null ? null : ampmRaw.Replace(".", "", StringComparison.Ordinal).ToUpperInvariant();

                // Build a timestamp string compatible with formats above
                string stamp = ampm is null || ampm.Length == 0 ? $"{d}, {t}" : $"{d}, {t} {ampm}";

                // First try invariant exact formats
                if (DateTime.TryParseExact(stamp, TimestampFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out timestamp))
                {
                    sender = m.Groups["name"].Value.Trim();
                    firstMessage = m.Groups["msg"].Value;
                    return true;
                }

                // Then try provided culture (ar-SA will accept Arabic context)
                if (DateTime.TryParse(stamp, culture, DateTimeStyles.None, out timestamp))
                {
                    sender = m.Groups["name"].Value.Trim();
                    firstMessage = m.Groups["msg"].Value;
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<ChatMessage> ParseChat(string path, CultureInfo culture)
        {
            using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);

            DateTime currentTs = default;
            string currentSender = "";
            var sb = new StringBuilder();
            string currentMedia = "";
            bool hasCurrent = false;
            bool currentSystem = false;

            string? line;
            while ((line = reader.ReadLine()) is not null)
            {
                if (TryParseHeader(line, culture, out var ts, out var sender, out var first))
                {
                    if (hasCurrent)
                    {
                        yield return new ChatMessage
                        {
                            Date = currentTs,
                            Sender = currentSender,
                            Message = sb.ToString().TrimEnd(),
                            Media = currentMedia,
                            IsSystem = currentSystem
                        };
                    }

                    currentTs = ts;
                    currentSender = sender;
                    sb.Clear();
                    sb.Append(first);
                    currentMedia = DetectMedia(first);
                    currentSystem = LooksSystemMessage(first);
                    hasCurrent = true;
                }
                else
                {
                    if (hasCurrent)
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
            }

            if (hasCurrent)
            {
                yield return new ChatMessage
                {
                    Date = currentTs,
                    Sender = currentSender,
                    Message = sb.ToString().TrimEnd(),
                    Media = currentMedia,
                    IsSystem = currentSystem
                };
            }
        }

        private static DateTimeOffset ApplyTimezone(DateTime dt, TimeSpan? tzOffset)
        {
            if (tzOffset is null)
                return new DateTimeOffset(dt); // as-is (assume local)
            // Treat parsed time as "wall clock" in that TZ, then convert to that offset
            return new DateTimeOffset(DateTime.SpecifyKind(dt, DateTimeKind.Unspecified), tzOffset.Value);
        }

        // === Excel writing (streaming) =======================================
        private static void WriteExcelStreaming(string inputPath, string outputPath, CultureInfo culture)
        {
            using var wb = new XLWorkbook();
            var sheets = new Dictionary<DateTime, (IXLWorksheet ws, int nextRow, int arabicScore)>();

            void EnsureSheet(DateTime day)
            {
                day = day.Date;
                if (sheets.ContainsKey(day)) return;

                var ws = wb.Worksheets.Add(day.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                ws.Cell(1, 1).Value = "Date";
                ws.Cell(1, 2).Value = "Sender";
                ws.Cell(1, 3).Value = "Message";
                ws.Cell(1, 4).Value = "Media";
                ws.Range(1, 1, 1, 4).Style.Font.Bold = true;
                ws.SheetView.FreezeRows(1);

                // widths
                ws.Column(1).Width = 20;
                ws.Column(2).Width = 28;
                ws.Column(3).Width = 90;
                ws.Column(4).Width = 30;

                sheets[day] = (ws, 2, 0);
            }

            foreach (var m in ParseChat(inputPath, culture))
            {
                if (_skipSystem && m.IsSystem)
                    continue;

                var dto = ApplyTimezone(m.Date, _tzOffset);
                var day = dto.Date;

                EnsureSheet(day);

                var entry = sheets[day];
                var ws = entry.ws;
                int row = entry.nextRow;

                ws.Cell(row, 1).Value = dto.LocalDateTime; // display local in the chosen TZ
                ws.Cell(row, 1).Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
                ws.Cell(row, 2).Value = m.Sender;
                ws.Cell(row, 3).Value = m.Message;
                ws.Cell(row, 3).Style.Alignment.WrapText = true;

                if (!string.IsNullOrEmpty(m.Media) && !string.IsNullOrEmpty(_mediaDir))
                {
                    var link = ResolveMediaLink(_mediaDir!, m.Media);
                    if (link is not null)
                    {
                        var cell = ws.Cell(row, 4);
                        cell.Value = Path.GetFileName(link);

                        var hl = cell.CreateHyperlink();
                        hl.ExternalAddress = new Uri(link, UriKind.Absolute);
                    }
                    else
                    {
                        ws.Cell(row, 4).Value = m.Media;
                    }
                }
                else
                {
                    ws.Cell(row, 4).Value = m.Media;
                }

                // quick RTL heuristic
                int score = entry.arabicScore;
                if (ContainsArabic(m.Message) || ContainsArabic(m.Sender))
                    score++;

                sheets[day] = (ws, row + 1, score);
            }

            // finalize: enable RTL if appropriate (or forced)
            foreach (var kv in sheets)
            {
                var (ws, nextRow, score) = kv.Value;
                // If more than ~20 Arabic rows or explicitly requested, flip RTL
                if (_forceRtl || score >= 20)
                    ws.RightToLeft = true;
            }

            wb.SaveAs(outputPath);
        }

        private static string? ResolveMediaLink(string mediaDir, string mediaToken)
        {
            // Exact path
            var exact = Path.Combine(mediaDir, mediaToken);
            if (File.Exists(exact)) return exact;

            // WhatsApp sometimes renames files; try a fuzzy match within the media directory
            var nameNoExt = Path.GetFileNameWithoutExtension(mediaToken);
            var ext = Path.GetExtension(mediaToken);
            try
            {
                var match = Directory.EnumerateFiles(mediaDir, $"*{nameNoExt}*{ext}", SearchOption.AllDirectories)
                                     .OrderByDescending(File.GetCreationTimeUtc)
                                     .FirstOrDefault();
                return match;
            }
            catch
            {
                return null;
            }
        }

        // === Entry point =====================================================
        private static void ParseArgs(string[] args, out string input, out string output)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("  WhatsAppChatToExcel <input_chat.txt> <output.xlsx> [--media-dir=\"C:\\Export\\Media\"] [--skip-system] [--culture=ar-SA] [--timezone=+03:00] [--rtl]");
                Environment.Exit(1);
            }

            input = Path.GetFullPath(args[0]);
            output = Path.GetFullPath(args[1]);
            if (!File.Exists(input))
            {
                Console.Error.WriteLine($"Input not found: {input}");
                Environment.Exit(2);
            }

            foreach (var a in args.Skip(2))
            {
                if (a.StartsWith("--media-dir=", StringComparison.OrdinalIgnoreCase))
                {
                    _mediaDir = a.Split('=', 2)[1].Trim('"');
                }
                else if (a.Equals("--skip-system", StringComparison.OrdinalIgnoreCase))
                {
                    _skipSystem = true;
                }
                else if (a.StartsWith("--culture=", StringComparison.OrdinalIgnoreCase))
                {
                    _cultureName = a.Split('=', 2)[1];
                }
                else if (a.StartsWith("--timezone=", StringComparison.OrdinalIgnoreCase))
                {
                    var s = a.Split('=', 2)[1];
                    if (TimeSpan.TryParse(s, out var tz))
                        _tzOffset = tz;
                    else
                        Console.WriteLine($"Warn: cannot parse timezone '{s}', expected like +03:00");
                }
                else if (a.Equals("--rtl", StringComparison.OrdinalIgnoreCase))
                {
                    _forceRtl = true;
                }
                else
                {
                    Console.WriteLine($"Warn: unknown option '{a}'");
                }
            }
        }

        private static void Main(string[] args)
        {
            ParseArgs(args, out var input, out var output);

            var culture = CultureInfo.InvariantCulture;
            if (!string.IsNullOrWhiteSpace(_cultureName))
            {
                try { culture = CultureInfo.GetCultureInfo(_cultureName!); }
                catch { Console.WriteLine($"Warn: culture '{_cultureName}' not found, using InvariantCulture."); }
            }

            try
            {
                WriteExcelStreaming(input, output, culture);
                Console.WriteLine($"Excel written: {output}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Failed: " + ex);
                Environment.Exit(3);
            }
        }
    }

    // Small helper to search Arabic spans quickly without allocations
    internal static class SpanExtensions
    {
        public static int IndexOfAnyInRange(this ReadOnlySpan<char> span, char minInclusive, char maxInclusive)
        {
            for (int i = 0; i < span.Length; i++)
            {
                var c = span[i];
                if (c >= minInclusive && c <= maxInclusive) return i;
            }
            return -1;
        }
    }
}
