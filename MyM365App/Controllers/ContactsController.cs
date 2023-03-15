using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Models;
using MyM365App.Services;


namespace MyM365App.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ContactsController : ControllerBase
    {
        private readonly ILogger<ContactsController> _logger;
        private readonly GraphContactClient _graphContactClient;

        public ContactsController(ILogger<ContactsController> logger, GraphContactClient graphContactClient)
        {
            _logger = logger;
            _graphContactClient = graphContactClient;
        }

        [HttpGet]
        [Route("GetContacts")]
        public async Task<IActionResult> GetStatsAsync()
        {
            try
            {
                var myContacts = await _graphContactClient.GetMyContacts();

                return Ok(myContacts);

            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }

        }
    }
}
