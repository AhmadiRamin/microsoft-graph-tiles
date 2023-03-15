using Microsoft.Graph;
using MyM365App.Graph;
using System.Text.Json;
using System.Web.Http.OData.Query;

namespace MyM365App.Services
{
    public class GraphPlannerClient
	{
        private readonly ILogger<GraphPlannerClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;
        public GraphPlannerClient(ILogger<GraphPlannerClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<List<PlannerTask>> GetMyTasks()
        {

            var tasks = new List<PlannerTask>();
            try
            {
                var results = await _graphServiceClient.Me.Planner.Tasks.Request().Select("id,title,percentComplete,dueDateTime").Top(10).GetAsync();
				tasks = results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }
            return tasks;
        }

		public async Task<List<PlannerTask>> GetAllTasksAsync()
		{
			List<PlannerTask> tasks = new List<PlannerTask>();
			try
			{
				var results = await _graphServiceClient.Me.Planner.Tasks.Request().Select("title,percentComplete").GetAsync();
				var pageIterator = PageIterator<PlannerTask>
									.CreatePageIterator(
										_graphServiceClient,
										results,
										// Callback executed for each item in
										// the collection
										(item) =>
										{
											tasks.Add(item);
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
			return tasks;
		}

	}
}
