using Microsoft.Graph;
using MyM365App.Constants;
using MyM365App.Models;
using System.IO;

namespace MyM365App.Helpers
{
    public class DocumentMapper
    {
        public static List<IDocument> RecentDocumentMapper(List<DriveItem> documents)
        {
            List<IDocument> result = new List<IDocument>();
            foreach (var doc in documents)
            {                
                
                string wwwRootFolder = Path.GetFullPath("wwwroot");
                string documentType = $"{Path.GetExtension(doc.Name).TrimStart('.')}";
                string documentIcon = $"/img/{Path.GetExtension(doc.Name).TrimStart('.')}.svg";                
                if (!System.IO.File.Exists($"{wwwRootFolder}{documentIcon}"))
                {
                    documentIcon = "img/document.svg";
                }

                    
                result.Add(new IDocument
                {
                    Title = doc.Name,
                    URL = doc.WebUrl,
                    path = doc.RemoteItem != null ? doc.RemoteItem.WebDavUrl.ToString() : string.Empty,
                    Icon = doc.Name != null  ? IconHelper.GetDocumentIcons(documentType) : "img/document.svg"
                });
            }
            return result;
        }

        public static List<IUsedDocuments> RecentlyUsedDocumentMapper(List<UsedInsight> documents)
        {
            List<IUsedDocuments> result = new List<IUsedDocuments>();

            foreach (var doc in documents)
            {
                result.Add(new IUsedDocuments
                {
                    Title = doc.ResourceVisualization.Title,
                    Type = doc.ResourceVisualization.Type,
                    QueryType = "Used",
                    Icon = IconHelper.GetDocumentIcons(doc.ResourceVisualization.Type),
                    PreviewImageUrl = doc.ResourceVisualization.PreviewImageUrl != null ? doc.ResourceVisualization.PreviewImageUrl : DefaultImages.UsedDocuments,
                    PreviewText = doc.ResourceVisualization.PreviewText,
                    ContainerDisplayName = doc.ResourceVisualization.ContainerDisplayName,
                    ContainerType = doc.ResourceVisualization.ContainerType,
                    ContainerWebUrl = Utilities.GetLocalPath(doc.ResourceVisualization.ContainerWebUrl),
                    URL = doc.ResourceReference.WebUrl,
                    Id = doc.ResourceReference.Id,
                    Accessed = doc.LastUsed.LastAccessedDateTime.HasValue ? doc.LastUsed.LastAccessedDateTime.Value.Date.ToShortDateString() : "",
                    Modified = doc.LastUsed.LastModifiedDateTime.HasValue ? doc.LastUsed.LastModifiedDateTime.Value.Date.ToShortDateString() : "",                    
                    LastModifiedSixMonthAgo = doc.LastUsed.LastModifiedDateTime.HasValue ? doc.LastUsed.LastModifiedDateTime.Value.Date > DateTime.Now.AddDays(-180) : false
                });
            }


            return result;
        }

        public static List<IOneDriveItems> OneDriveItemsMapper(List<DriveItem> documents)
        {
            List<IOneDriveItems> result = new List<IOneDriveItems>();
            foreach (var doc in documents)
            {
                                
                result.Add(new IOneDriveItems
                {
                    Title = doc.Name,
                    URL = doc.WebUrl,
                    path = doc.RemoteItem != null ? doc.RemoteItem.WebDavUrl.ToString() : string.Empty,
                    Icon = doc.Folder != null ? IconHelper.GetDocumentIcons("folder") : IconHelper.GetDocumentIcons(doc.Name != null ? doc.Name : "Document"),
                    Modified = doc.LastModifiedDateTime.HasValue ? doc.LastModifiedDateTime.Value.Date.ToShortDateString() : "",
                }) ;
            }
            return result;
        }


    }
}
