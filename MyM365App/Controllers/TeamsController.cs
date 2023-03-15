using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Models;
using MyM365App.Services;


namespace MyM365App.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TeamsController : ControllerBase
    {
        private readonly ILogger<TeamsController> _logger;
        private readonly GraphTeamClient _graphTeamClient;

        public TeamsController(ILogger<TeamsController> logger, GraphTeamClient graphTeamClient)
        {
            _logger = logger;
            _graphTeamClient = graphTeamClient;
        }

        [HttpGet]
        [Route("GetStats")]
        public async Task<IActionResult> GetStatsAsync()
        {
            try
            {
                var installedAppsCount = await _graphTeamClient.GetInstalledAppsCount();
                var joinedTeams = await _graphTeamClient.GetAssociatedTeamsCount();

                var stats = new TeamsStats
                {
                    InstalledAppsCount = installedAppsCount,
                    AssociatedTeamsCount = joinedTeams
                };

                return Ok(stats);

            }
            catch (ServiceException ex)
            {
                _logger.LogError(ex.Message);
                return BadRequest(ex.Message);
            }

        }

        [HttpGet]
        [Route("GetTeams")]
        public async Task<IActionResult> GetTeams()
        {
			try
			{
				var teams = await _graphTeamClient.GetTeams();

				return Ok(teams.Take(10));

			}
			catch (ServiceException ex)
			{
				_logger.LogError(ex.Message);
				return BadRequest(ex.Message);
			}
		}


    }
}
