using ClosedXML.Excel;
using System;
using System.Linq;

namespace WhatsAppChatToExcel
{
    internal sealed class ExcelWriter
    {
        private readonly CommandLineOptions _options;
        private readonly ChatParserOptions _parserOptions;

        public ExcelWriter(CommandLineOptions options, ChatParserOptions parserOptions)
        {
            _options = options;
            _parserOptions = parserOptions;
        }

        public void WriteExcel(IEnumerable<ChatMessage> messages, string outputPath)
        {
            if (_options.SheetMode == SheetMode.All)
                WriteAllInOneSheet(messages, outputPath);
            else
                WritePerDay(messages, outputPath);
        }

        private void WritePerDay(IEnumerable<ChatMessage> messages, string outputPath)
        {
            using var wb = new XLWorkbook();
            var sheets = new Dictionary<DateTime, (IXLWorksheet ws, int nextRow, int arabicScore)>();

            void EnsureSheet(DateTime day)
            {
                if (sheets.ContainsKey(day)) return;
                var ws = wb.Worksheets.Add(day.ToString("yyyy-MM-dd"));
                SetupWorksheet(ws);
                sheets[day] = (ws, 2, 0);
            }

            foreach (var message in GetFilteredMessages(messages))
            {
                var dto = ApplyTimezone(message.Date, _parserOptions.TimezoneOffset);
                var day = dto.Date;

                EnsureSheet(day);
                var (ws, row, score) = sheets[day];

                WriteMessageToRow(ws, row, message, dto);

                if (ContainsArabic(message.Message) || ContainsArabic(message.Sender)) score++;
                sheets[day] = (ws, row + 1, score);
            }

            foreach (var kv in sheets)
                if (_options.ForceRtl || kv.Value.arabicScore >= 20)
                    kv.Value.ws.RightToLeft = true;

            wb.SaveAs(outputPath);
        }

        private void WriteAllInOneSheet(IEnumerable<ChatMessage> messages, string outputPath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("All");
            SetupWorksheet(ws);

            int row = 2;
            int arabicScore = 0;

            foreach (var message in GetFilteredMessages(messages))
            {
                var dto = ApplyTimezone(message.Date, _parserOptions.TimezoneOffset);
                WriteMessageToRow(ws, row, message, dto);

                if (ContainsArabic(message.Message) || ContainsArabic(message.Sender)) arabicScore++;
                row++;
            }

            if (_options.ForceRtl || arabicScore >= 20)
                ws.RightToLeft = true;

            wb.SaveAs(outputPath);
        }

        private static void SetupWorksheet(IXLWorksheet ws)
        {
            ws.Cell(1, 1).Value = "Date";
            ws.Cell(1, 2).Value = "Sender";
            ws.Cell(1, 3).Value = "Message";
            ws.Range(1, 1, 1, 3).Style.Font.Bold = true;
            ws.SheetView.FreezeRows(1);
            ws.Column(1).Width = 20;
            ws.Column(2).Width = 28;
            ws.Column(3).Width = 90;
        }

        private void WriteMessageToRow(IXLWorksheet ws, int row, ChatMessage message, DateTimeOffset dto)
        {
            ws.Cell(row, 1).Value = dto.LocalDateTime;
            ws.Cell(row, 1).Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
            ws.Cell(row, 2).Value = message.Sender;
            ws.Cell(row, 3).Value = message.Message;
            ws.Cell(row, 3).Style.Alignment.WrapText = true;

            if (!string.IsNullOrEmpty(_options.MediaDirectory))
            {
                var token = MediaHelper.DetectMediaToken(message.Message);
                if (!string.IsNullOrEmpty(token))
                {
                    var link = MediaHelper.ResolveMediaLink(_options.MediaDirectory!, token);
                    if (link != null)
                    {
                        var msgCell = ws.Cell(row, 3);
                        var hl = msgCell.CreateHyperlink();
                        hl.ExternalAddress = new Uri(link, UriKind.Absolute);
                    }
                }
            }
        }

        private IEnumerable<ChatMessage> GetFilteredMessages(IEnumerable<ChatMessage> messages)
        {
            foreach (var message in messages)
            {
                if (_options.SkipSystem && message.IsSystem) continue;

                var dto = ApplyTimezone(message.Date, _parserOptions.TimezoneOffset);
                var day = dto.Date;

                if (_options.FromDate.HasValue && day < _options.FromDate.Value) continue;
                if (_options.ToDate.HasValue && day > _options.ToDate.Value) continue;

                yield return message;
            }
        }

        private static DateTimeOffset ApplyTimezone(DateTime dt, TimeSpan? tzOffset) =>
            tzOffset == null ? new DateTimeOffset(dt) : new DateTimeOffset(DateTime.SpecifyKind(dt, DateTimeKind.Unspecified), tzOffset.Value);

        private static bool ContainsArabic(string s) =>
            s.AsSpan().IndexOfAnyInRange('\u0600', '\u06FF') >= 0
            || s.AsSpan().IndexOfAnyInRange('\u0750', '\u077F') >= 0
            || s.AsSpan().IndexOfAnyInRange('\u08A0', '\u08FF') >= 0;
    }
}
