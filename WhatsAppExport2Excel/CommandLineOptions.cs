using System;
using System.Globalization;
using System.Linq;

namespace WhatsAppChatToExcel;

internal sealed class CommandLineOptions
{
    public string InputPath { get; init; } = "";
    public string OutputPath { get; init; } = "";
    public string? MediaDirectory { get; init; }
    public bool SkipSystem { get; init; }
    public string? CultureName { get; init; }
    public TimeSpan? TimezoneOffset { get; init; }
    public bool ForceRtl { get; init; }
    public DateTime? FromDate { get; init; }
    public DateTime? ToDate { get; init; }
    public SheetMode SheetMode { get; init; } = SheetMode.Day;

    public static CommandLineOptions Parse(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage:\n  WhatsAppChatToExcel <input_chat.txt> <output.xlsx> [--media-dir=...] [--skip-system] [--culture=ar-SA] [--timezone=+03:00] [--rtl] [--from=YYYY-MM-DD] [--to=YYYY-MM-DD] [--sheet=day|all]");
            Environment.Exit(1);
        }

        var options = new CommandLineOptions
        {
            InputPath = Path.GetFullPath(args[0]),
            OutputPath = Path.GetFullPath(args[1])
        };

        string? mediaDir = null;
        bool skipSystem = false;
        string? cultureName = null;
        TimeSpan? tzOffset = null;
        bool forceRtl = false;
        DateTime? fromDate = null;
        DateTime? toDate = null;
        SheetMode sheetMode = SheetMode.Day;

        foreach (var arg in args.Skip(2))
        {
            if (arg.StartsWith("--media-dir=", StringComparison.OrdinalIgnoreCase))
                mediaDir = arg.Split('=', 2)[1].Trim('"');
            else if (arg.Equals("--skip-system", StringComparison.OrdinalIgnoreCase))
                skipSystem = true;
            else if (arg.StartsWith("--culture=", StringComparison.OrdinalIgnoreCase))
                cultureName = arg.Split('=', 2)[1];
            else if (arg.StartsWith("--timezone=", StringComparison.OrdinalIgnoreCase) &&
                     TimeSpan.TryParse(arg.Split('=', 2)[1], out var tz))
                tzOffset = tz;
            else if (arg.Equals("--rtl", StringComparison.OrdinalIgnoreCase))
                forceRtl = true;
            else if (arg.StartsWith("--from=", StringComparison.OrdinalIgnoreCase) &&
                     DateTime.TryParseExact(arg.Split('=', 2)[1], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d1))
                fromDate = d1.Date;
            else if (arg.StartsWith("--to=", StringComparison.OrdinalIgnoreCase) &&
                     DateTime.TryParseExact(arg.Split('=', 2)[1], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d2))
                toDate = d2.Date;
            else if (arg.StartsWith("--sheet=", StringComparison.OrdinalIgnoreCase))
            {
                var value = arg.Split('=', 2)[1].Trim().ToLowerInvariant();
                sheetMode = value == "all" ? SheetMode.All : SheetMode.Day;
            }
        }

        return new CommandLineOptions
        {
            InputPath = options.InputPath,
            OutputPath = options.OutputPath,
            MediaDirectory = mediaDir,
            SkipSystem = skipSystem,
            CultureName = cultureName,
            TimezoneOffset = tzOffset,
            ForceRtl = forceRtl,
            FromDate = fromDate,
            ToDate = toDate,
            SheetMode = sheetMode
        };
    }
}
