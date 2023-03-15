using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Graph;
using MyM365App.Models;
using Newtonsoft.Json;
using static MyM365App.Controllers.ProxyController;
using System.Net;
using System.Net.Http.Headers;
using System.Text;

namespace MyM365App.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class MessagesController : ControllerBase
	{
		private readonly ILogger<GraphProfileClient> _logger;
		private readonly GraphMessageClient _graphMessageClient;

		public MessagesController(ILogger<GraphProfileClient> logger, GraphMessageClient graphMessageClient)
		{
			_logger = logger;
			_graphMessageClient = graphMessageClient;
		}

		[HttpGet]
		[Route("GetStats")]
		public async Task<IActionResult> GetStatsAsync(string userEmail)
		{
			try
			{
				var unReadMessagesCount = await _graphMessageClient.GetMessagesCountAsync("isRead eq false");
				var readMessagesCount = await _graphMessageClient.GetMessagesCountAsync("isRead eq true");
				var sentMessagesCount = await _graphMessageClient.GetMessagesCountAsync($"from/emailAddress/address eq '{userEmail}'");
				var receivedMessagesCount = await _graphMessageClient.GetMessagesCountBySearchAsync($"%22to:{userEmail}%22");
				var draftMessagesCount = await _graphMessageClient.GetMessagesCountAsync("isDraft eq true");
				var stats = new MessageStats
				{
					TotalReadMessagesCount = readMessagesCount,
					TotalUnreadMessagesCount = unReadMessagesCount,
					TotalSentMessagesCount = sentMessagesCount,
					TotalReceivedMessagesCount = receivedMessagesCount,
					TotalDraftMessagesCount = draftMessagesCount
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
