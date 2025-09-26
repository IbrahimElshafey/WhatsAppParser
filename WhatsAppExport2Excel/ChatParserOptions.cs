using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace WhatsAppChatToExcel;

internal sealed class ChatParserOptions
{
    public static readonly Regex[] LineStartPatterns =
    {
        new Regex(@"^(?<d>\d{1,2}/\d{1,2}/\d{2,4}),\s*(?<t>\d{1,2}:\d{2}(?::\d{2})?)\s*(?<ampm>(?:[AaPp]\.?[Mm]\.?)?)\s*[-–]\s*(?<name>.+?):\s*(?<msg>.*)$",
            RegexOptions.Compiled),
        new Regex(@"^\[\s*(?<d>\d{1,2}/\d{1,2}/\d{2,4}),\s*(?<t>\d{1,2}:\d{2}(?::\d{2})?)\s*(?<ampm>(?:[AaPp]\.?[Mm]\.?)?)\s*\]\s*(?<name>.+?):\s*(?<msg>.*)$",
            RegexOptions.Compiled),
    };

    public static readonly string[] TimestampFormats =
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

    public static readonly string[] SystemMessageStarts =
    {
        "messages to this chat are now",
        "security code changed",
        "missed voice call", "missed video call",
        "تم إنشاء المجموعة", "قام بتغيير صورة المجموعة", "قام بتغيير وصف المجموعة",
        "تم تغيير رقم الهاتف", "أصبحت الرسائل الآن"
    };

    public static readonly Regex FileLikePattern =
        new Regex(@"\b[\p{L}\p{Nd}_\-]+\.(jpg|jpeg|png|gif|mp4|mp3|opus|pdf|docx?|xlsx?|pptx?|heic|mov|zip)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;
    public bool SkipSystemMessages { get; init; }
    public TimeSpan? TimezoneOffset { get; init; }
    public DateTime? FromDate { get; init; }
    public DateTime? ToDate { get; init; }
    public string? MediaDirectory { get; init; }
}
