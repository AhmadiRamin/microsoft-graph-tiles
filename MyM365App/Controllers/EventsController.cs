using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Services;


namespace MyM365App.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EventsController : ControllerBase
    {
        private readonly ILogger<EventsController> _logger;
        private readonly GraphCalendarClient _graphCalendarClient;

        public EventsController(ILogger<EventsController> logger, GraphCalendarClient graphCalendarClient)
        {
            _logger = logger;
            _graphCalendarClient = graphCalendarClient;
        }

        [HttpGet]
        [Route("GetEvents")]
        public async Task<IActionResult> GetEventsAsync()
        {
            try
            {
                var myContacts = await _graphCalendarClient.GetMyEvents();

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
