using System;
using System.Linq;

namespace WhatsAppChatToExcel;

internal sealed class ChatMessage
{
    public DateTime Date { get; set; }
    public string Sender { get; set; } = "";
    public string Message { get; set; } = "";
    public bool IsSystem { get; set; }
}
