using Microsoft.Graph;
using MyM365App.Helpers;
using MyM365App.Models;

namespace MyM365App.Mapper
{
    public class NotebookMapper
    {
        public static List<INotebook> MyNotebooksMapper(List<Notebook> notebooks)
        {
            List<INotebook> result = new List<INotebook>();
            foreach (var notebook in notebooks)
            {
                               

                result.Add(new INotebook
                {
                    DisplayName = notebook.DisplayName,
                    Icon = IconHelper.GetDocumentIcons("onenote"),
                    Links = notebook.Links,
                    LastModified = notebook.LastModifiedDateTime.HasValue ? notebook.LastModifiedDateTime.Value.Date.ToShortDateString() : "",
                    LastModifiedBy = notebook.LastModifiedBy != null ? notebook.LastModifiedBy.User.DisplayName : string.Empty
                });
            }
            return result;
        }
    }
}
