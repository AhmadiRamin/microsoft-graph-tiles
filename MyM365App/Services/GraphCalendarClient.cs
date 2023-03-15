using Microsoft.Graph;
using MyM365App.Graph;
using System.Text.Json;
using System.Web.Http.OData.Query;

namespace MyM365App.Services
{
	public class GraphCalendarClient
	{
		private readonly ILogger<GraphCalendarClient> _logger;
		private readonly GraphServiceClient _graphServiceClient;
		public GraphCalendarClient(ILogger<GraphCalendarClient> logger, GraphServiceClient graphServiceClient)
		{
			_logger = logger;
			_graphServiceClient = graphServiceClient;
		}

		public async Task<List<Event>> GetMyEvents()
		{

			var events = new List<Event>();
			try
			{
				var results = await _graphServiceClient.Me.Events.Request().Top(10).Select("webLink, subject").GetAsync();
				events = results.ToList();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message);
			}
			return events;
		}

	}
}
