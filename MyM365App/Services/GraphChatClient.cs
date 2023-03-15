using Microsoft.Graph;
using MyM365App.Graph;
using System.Text.Json;
using System.Web.Http.OData.Query;

namespace MyM365App.Services
{
    public class GraphChatClient
    {
        private readonly ILogger<GraphDriveClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;
        public GraphChatClient(ILogger<GraphDriveClient> logger, GraphServiceClient graphServiceClient) {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<List<Chat>> GetMyChats()
        {
            
            var chats = new List<Chat>();
            try
            {
                var results = await _graphServiceClient.Me.Chats.Request().GetAsync();
                var pageIterator = PageIterator<Chat>
                    .CreatePageIterator(
                        _graphServiceClient,
                        results,
                        // Callback executed for each item in
                        // the collection
                        (item) =>
                        {
                            chats.Add(item);
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
            catch(Exception ex)
            {
                _logger.LogError(ex.Message);
            }
            return chats;
        }

        public async Task<Chat> GetLastPersonIChatWith()
        {
            var lastChat = new Chat();
            try
            {
                var results = await _graphServiceClient.Me.Chats.Request().Filter("chatType eq 'OneOnOne'").Top(1).Expand("members").GetAsync();
                if (results != null)
                    lastChat = results.First();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }
            return lastChat;
        }
    }
}
