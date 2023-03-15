using Microsoft.Graph;

namespace MyM365App.Models
{
    public class INotebook
    {
        public string? DisplayName { get; set; }
        public NotebookLinks? Links { get; set; }

        public string? Icon { get; set; }
        public string? LastModified { get; set;}

        public string? LastModifiedBy { get; set; }
    }
}
