using Microsoft.Graph;
using MyM365App.Graph;
using System.Text.Json;
using System.Web.Http.OData.Query;

namespace MyM365App.Services
{
    public class GraphContactClient
    {
        private readonly ILogger<GraphContactClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;
        public GraphContactClient(ILogger<GraphContactClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<List<Contact>> GetMyContacts()
        {

            var contacts = new List<Contact>();
            try
            {
                var results = await _graphServiceClient.Me.Contacts.Request().Select("displayName, emailAddresses").Top(10).GetAsync();
                contacts = results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }
            return contacts;
        }

    }
}
