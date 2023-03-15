using Microsoft.Graph;
using MyM365App.Models;

namespace MyM365App.ViewModels
{
    public class ViewModel
    {
        public List<IDocument> RecentDocuments { get; set; }
        public List<IUsedDocuments> RecentlyUsedDocuments { get; set; }

        public List<IOneDriveItems> MyOneDriveItems { get; set; }

        public List<INotebook> MyNotebooks { get; set; }

    }
}
