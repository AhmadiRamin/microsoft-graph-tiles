using Microsoft.Extensions.Options;
using Microsoft.Graph;
using MyM365App.Graph;
using System.Text.Json;
using System.Web.Http.OData.Query;

namespace MyM365App.Services
{
    public class GraphTeamClient
    {
        private readonly ILogger<GraphTeamClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;
        public GraphTeamClient(ILogger<GraphTeamClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<Int32> GetInstalledAppsCount()
        {
            var count = 0;
            try
            {
                var results = await _graphServiceClient.Me.Teamwork.InstalledApps.Request().GetAsync();

                if (results.AdditionalData.TryGetValue("@odata.count", out var propertyValue))
                {
                    var propertyElement = (JsonElement)propertyValue;
                    count = propertyElement.Deserialize<Int32>();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }
            return count;
        }

        public async Task<Int32> GetAssociatedTeamsCount()
        {
            var count = 0;
            try
            {
                var results = await _graphServiceClient.Me.JoinedTeams.Request().GetAsync();

                if (results.AdditionalData.TryGetValue("@odata.count", out var propertyValue))
                {
                    var propertyElement = (JsonElement)propertyValue;
                    count = propertyElement.Deserialize<Int32>();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }
            return count;
        }

        public async Task<List<Team>> GetTeams()
        {
			var teams = new List<Team>();
			try
			{
				var results = await _graphServiceClient.Me.JoinedTeams.Request().GetAsync();
				teams = results.ToList();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message);
			}
			return teams;
		}

    }
}
