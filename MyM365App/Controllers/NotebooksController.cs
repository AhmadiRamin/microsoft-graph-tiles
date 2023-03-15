using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using MyM365App.Graph;
using MyM365App.Models;
using Newtonsoft.Json;
using static MyM365App.Controllers.ProxyController;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using MyM365App.Helpers;
using MyM365App.Services;
using MyM365App.Mapper;

namespace MyM365App.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class NotebooksController : ControllerBase
	{
		private readonly ILogger<GraphProfileClient> _logger;
		private readonly GraphOneNoteClient _graphOneNoteClient;

		public NotebooksController(ILogger<GraphProfileClient> logger, GraphOneNoteClient graphOneNoteClient)
		{
			_logger = logger;
			_graphOneNoteClient = graphOneNoteClient;
		}

		[HttpGet]
		[Route("GetNotebooks")]
		public async Task<IActionResult> GetNotebooksAsync()
        {
            try
			{
				var notebooks = await _graphOneNoteClient.GetNoteBooksAsync();
				List<INotebook> response = NotebookMapper.MyNotebooksMapper(notebooks);
				return Ok(response);

			}
			catch (ServiceException ex)
			{
				return BadRequest(ex.Message);
			}

		}

       
    }
}
