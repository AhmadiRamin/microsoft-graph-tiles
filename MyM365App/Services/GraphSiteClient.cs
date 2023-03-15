using Microsoft.Graph;
using MyM365App.Graph;
using System.Text.Json;
using System.Web.Http.OData.Query;

namespace MyM365App.Services
{
	public class GraphSiteClient
	{
		private readonly ILogger<GraphSiteClient> _logger;
		private readonly GraphServiceClient _graphServiceClient;
		public GraphSiteClient(ILogger<GraphSiteClient> logger, GraphServiceClient graphServiceClient)
		{
			_logger = logger;
			_graphServiceClient = graphServiceClient;
		}

		public async Task<List<Site>> GetFollowedSites()
		{

			var sites = new List<Site>();
			try
			{
				var results = await _graphServiceClient.Me.FollowedSites.Request().Select("webUrl, displayName").Top(10).GetAsync();
				sites = results.ToList();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message);
			}
			return sites;
		}

	}
}
