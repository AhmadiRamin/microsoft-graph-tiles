using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Models;
using MyM365App.Services;


namespace MyM365App.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PeopleController : ControllerBase
    {
        private readonly ILogger<PeopleController> _logger;
        private readonly GraphPeopleClient _graphPeopleClient;

        public PeopleController(ILogger<PeopleController> logger, GraphPeopleClient graphPeopleClient)
        {
            _logger = logger;
			_graphPeopleClient = graphPeopleClient;
        }

        [HttpGet]
        [Route("GetColleagues")]
        public async Task<IActionResult> GetColleaguesAsync()
        {
            try
            {
                var myContacts = await _graphPeopleClient.GetMyColleagues();

                return Ok(myContacts);

            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }

        }
    }
}
