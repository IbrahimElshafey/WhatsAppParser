using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace WhatsAppChatToExcel;
internal static class Program
{
    //"E:\SourceCode\Mine\WhatsApp Chat with DIYAR-KFIA\WhatsApp Chat with DIYAR-KFIA.txt" "E:\SourceCode\Mine\WhatsApp Chat with DIYAR-KFIA\Chat-all123.xlsx" --media-dir="E:\SourceCode\Mine\WhatsApp Chat with DIYAR-KFIA" --culture=ar-SA --timezone=+03:00 --skip-system --from=2025-09-04 --to=2025-09-22 --sheet=all
    private static void Main(string[] args)
    {
        var parserOptions = ChatParserOptions.LoadFromSettingsFile();
        var parser = new ChatParser(parserOptions);
        var messages = parser.ParseChat();

        var excelWriter = new ExcelWriter(parserOptions);
        excelWriter.WriteExcel(messages, parserOptions.OutputPath);

        var includedMedia = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Console.WriteLine($"\nExcel written: {parserOptions.OutputPath}");
        if (parserOptions.FromDate.HasValue || parserOptions.ToDate.HasValue)
            Console.WriteLine($"Filtered days: {(parserOptions.FromDate?.ToString("yyyy-MM-dd") ?? "min")} → {(parserOptions.ToDate?.ToString("yyyy-MM-dd") ?? "max")}");
        Console.WriteLine($"Sheet mode: {parserOptions.SheetMode}");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
