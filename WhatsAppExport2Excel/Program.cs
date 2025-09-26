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
    private static void Main(string[] args)
    {
        var cmdOptions = CommandLineOptions.Parse(args);

        var culture = CultureInfo.InvariantCulture;
        if (!string.IsNullOrWhiteSpace(cmdOptions.CultureName))
        {
            try { culture = CultureInfo.GetCultureInfo(cmdOptions.CultureName); }
            catch { /* Use default culture */ }
        }

        var parserOptions = new ChatParserOptions
        {
            Culture = culture,
            SkipSystemMessages = cmdOptions.SkipSystem,
            TimezoneOffset = cmdOptions.TimezoneOffset,
            FromDate = cmdOptions.FromDate,
            ToDate = cmdOptions.ToDate,
            MediaDirectory = cmdOptions.MediaDirectory
        };

        var parser = new ChatParser(parserOptions);
        var messages = parser.ParseChat(cmdOptions.InputPath);

        var excelWriter = new ExcelWriter(cmdOptions, parserOptions);
        excelWriter.WriteExcel(messages, cmdOptions.OutputPath);

        Console.WriteLine($"\nExcel written: {cmdOptions.OutputPath}");
        if (cmdOptions.FromDate.HasValue || cmdOptions.ToDate.HasValue)
            Console.WriteLine($"Filtered days: {(cmdOptions.FromDate?.ToString("yyyy-MM-dd") ?? "min")} → {(cmdOptions.ToDate?.ToString("yyyy-MM-dd") ?? "max")}");
        Console.WriteLine($"Sheet mode: {cmdOptions.SheetMode}");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
