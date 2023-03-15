namespace MyM365App.Models
{
    public class IUsedDocuments
    {
        public string? Title { get; set; }
        public string? Type { get; set; }
        public string? Icon { get; set; }
        public string? PreviewImageUrl { get; set; }
        public string? PreviewText { get; set; }
        public string? ContainerDisplayName { get; set; }
        public string? ContainerType { get; set; }
        public string? ContainerWebUrl { get; set; }
        public string? URL { get; set; }
        public string? Accessed { get; set; }
        public string? Modified { get; set; }        
        public string? QueryType { get; set; }
        public bool LastModifiedSixMonthAgo { get; set; }
        public string? Id { get; set; }
    }
}
