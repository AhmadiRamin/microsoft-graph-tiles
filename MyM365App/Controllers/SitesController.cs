using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Services;


namespace MyM365App.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SitesController : ControllerBase
    {
        private readonly ILogger<SitesController> _logger;
        private readonly GraphSiteClient _graphSiteClient;

        public SitesController(ILogger<SitesController> logger, GraphSiteClient graphSiteClient)
        {
            _logger = logger;
			_graphSiteClient = graphSiteClient;
        }

        [HttpGet]
        [Route("GetSites")]
        public async Task<IActionResult> GetSitesAsync()
        {
            try
            {
                var myContacts = await _graphSiteClient.GetFollowedSites();

                return Ok(myContacts);

            }
            catch (ServiceException ex)
            {
                _logger.LogError(ex.Message);
                return BadRequest(ex.Message);
            }

        }
    }
}
