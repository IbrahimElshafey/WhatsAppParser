using ClosedXML.Excel;
using System;
using System.Linq;

namespace WhatsAppChatToExcel;

internal static class SpanExtensions
{
    public static int IndexOfAnyInRange(this ReadOnlySpan<char> span, char minInclusive, char maxInclusive)
    {
        for (int i = 0; i < span.Length; i++)
            if (span[i] >= minInclusive && span[i] <= maxInclusive) return i;
        return -1;
    }
}
