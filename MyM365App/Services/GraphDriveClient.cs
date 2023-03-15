using Microsoft.Graph;
using MyM365App.Models;

namespace MyM365App.Graph
{
    public class GraphDriveClient
    {
        private readonly ILogger<GraphDriveClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;
        public GraphDriveClient(ILogger<GraphDriveClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<List<DriveItem>> GetFiles()
        {
            List<DriveItem> items = new List<DriveItem>();
            try
            {
                var results = await _graphServiceClient.Me.Drive.Root.Search("").Request().GetAsync();
                
                var pageIterator = PageIterator<DriveItem>
                    .CreatePageIterator(
                        _graphServiceClient,
                        results,
                        // Callback executed for each item in
                        // the collection
                        (item) =>
                        {
                            items.Add(item);
                            return true;
                        },
                        // Used to configure subsequent page
                        // requests
                        (req) =>
                        {
                            // Re-add the header to subsequent requests
                            //req.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                            return req;
                        }
                    );

                await pageIterator.IterateAsync();
            }
            catch (Exception ex)
            {

            }
            return items;
        }

        public async Task<List<DriveItem>> GetRecentDocumentsAsync()
        {
            List<DriveItem> items = new List<DriveItem>();
            try
            {
                var selectedFields = new string[] { "name", "webUrl", "lastModifiedDateTime", "lastModifiedBy", "remoteItem" };

                var results = await _graphServiceClient.Me.Drive
                                                         .Recent()
                                                         .Request()
                                                         .Top(10)
                                                         .Select("name, webUrl")
                                                         .GetAsync();
                items = results.ToList();
            }
            catch (Exception ex)
            {

            }
            return items;
        }

        public async Task<List<UsedInsight>> GetUsedDocumentsAsync()
        {
            List<UsedInsight> items = new List<UsedInsight>();
            try
            {
                
                var results = await _graphServiceClient.Me.Insights
                                                         .Used
                                                         .Request()
                                                         .Top(10)
                                                         .GetAsync();

                items =  results.ToList();

            }
            catch (Exception ex)
            {
                                
            }

            return items;
            
        }

        public async Task<List<DriveItem>> GetOneDriveItemsAsync()
        {
            List<DriveItem> items = new List<DriveItem>();
            try
            {
                var selectedFields = new string[] { "name", "webUrl", "lastModifiedDateTime", "lastModifiedBy", "remoteItem", "file","folder" };

                var results = await _graphServiceClient.Me.Drive
                                                         .Root
                                                         .Children
                                                         .Request()
                                                         .Top(10)
                                                         .Select(String.Join(",", selectedFields))
                                                         .GetAsync();

               items = results.ToList();

            }
            catch (Exception ex)
            {

            }
            return items;
        }
    }
}

