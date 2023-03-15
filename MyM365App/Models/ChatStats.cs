using Microsoft.Graph;

namespace MyM365App.Models
{
    public class ChatStats
    {
        public int TotalGroupChatsCount { get; set; }
        public int TotalOneOnOneChatsCount { get; set; }
        public int TotalMeetingChatsCount { get; set; }
        public Chat MyLastChat { get; set; }
    }
}
