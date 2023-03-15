using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Models;
using MyM365App.Services;


namespace MyM365App.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ChatsController : ControllerBase
    {
        private readonly ILogger<ChatsController> _logger;
        private readonly GraphChatClient _graphChatClient;

        public ChatsController(ILogger<ChatsController> logger, GraphChatClient graphChatClient)
        {
            _logger = logger;
            _graphChatClient = graphChatClient;
        }

        [HttpGet]
        [Route("GetStats")]
        public async Task<IActionResult> GetStatsAsync()
        {
            try
            {
                var myChats = await _graphChatClient.GetMyChats();
                var myLastChat = await _graphChatClient.GetLastPersonIChatWith();
                
                var stats = new ChatStats
                {
                    TotalGroupChatsCount = myChats.Count(c => c.ChatType == ChatType.Group),
                    TotalMeetingChatsCount = myChats.Count(c => c.ChatType == ChatType.Meeting),
                    TotalOneOnOneChatsCount = myChats.Count(c => c.ChatType == ChatType.OneOnOne),
                    MyLastChat = myLastChat
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
