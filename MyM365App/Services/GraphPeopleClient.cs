using Microsoft.Graph;
using MyM365App.Graph;
using System.Text.Json;
using System.Web.Http.OData.Query;

namespace MyM365App.Services
{
    public class GraphPeopleClient
	{
        private readonly ILogger<GraphPeopleClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;
        public GraphPeopleClient(ILogger<GraphPeopleClient> logger, GraphServiceClient graphServiceClient) {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<List<Person>> GetMyColleagues()
        {
            
            var people = new List<Person>();
            try
            {
                var results = await _graphServiceClient.Me.People
                    .Request()
                    .Filter("personType/class eq 'Person' and personType/subclass eq 'OrganizationUser'")
                    .Select("displayName,jobTitle,userPrincipalName,imAddress,personType")
                    .GetAsync();
                people = results.ToList();
            }
            catch(Exception ex)
            {
                _logger.LogError(ex.Message);
            }
            return people;
        }

    }
}
