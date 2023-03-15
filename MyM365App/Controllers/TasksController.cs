using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Models;
using MyM365App.Services;


namespace MyM365App.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TasksController : ControllerBase
    {
        private readonly ILogger<TasksController> _logger;
        private readonly GraphPlannerClient _graphPlannerClient;

        public TasksController(ILogger<TasksController> logger, GraphPlannerClient graphPlannerClient)
        {
            _logger = logger;
			_graphPlannerClient = graphPlannerClient;
        }

        [HttpGet]
        [Route("GetTasks")]
        public async Task<IActionResult> GetTasksAsync()
        {
            try
            {
                var tasks = await _graphPlannerClient.GetMyTasks();

                return Ok(tasks);

            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }

        }

		[HttpGet]
		[Route("GetTaskStats")]
		public async Task<IActionResult> GetTaskStatsAsync()
		{
			try
			{
				var tasks = await _graphPlannerClient.GetAllTasksAsync();

                var stats = new TaskStats
                {
                    TotalCompletedTasksCount = tasks.Count(a=> a.PercentComplete == 100),
                    TotalInprogressTasksCount = tasks.Count(a => a.PercentComplete > 0 && a.PercentComplete < 100),
					TotalNotStartedTasksCount = tasks.Count(a => a.PercentComplete == 0),
				};
				return Ok(stats);
			}
			catch (ServiceException ex)
			{
				return BadRequest(ex.Message);
			}

		}

	}
}
