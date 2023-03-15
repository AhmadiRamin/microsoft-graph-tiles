namespace MyM365App.Models
{
    public class MessageStats
    {
        public int TotalMessagesCount { get; set;}
        public int TotalReceivedMessagesCount { get;set; }
        public int TotalSentMessagesCount { get; set; }
        public int TotalDraftMessagesCount { get; set; }
        public int TotalReadMessagesCount { get; set; }
        public int TotalUnreadMessagesCount { get; set; }
    }
}
