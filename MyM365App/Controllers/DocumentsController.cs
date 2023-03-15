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
using System.Text.RegularExpressions;

namespace MyM365App.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class DocumentsController : ControllerBase
	{
		private readonly ILogger<GraphProfileClient> _logger;
		private readonly GraphDriveClient _graphDriveClient;

		public DocumentsController(ILogger<GraphProfileClient> logger, GraphDriveClient graphDriveClient)
		{
			_logger = logger;
            _graphDriveClient = graphDriveClient;
		}

		[HttpGet]
		[Route("GetRecentDocuments")]
		public async Task<IActionResult> GetRecentDocumentsAsync()
        {
            try
			{
                var files = await _graphDriveClient.GetRecentDocumentsAsync();                
                List<IDocument> response = DocumentMapper.RecentDocumentMapper(files.Cast<DriveItem>().ToList());                
				return Ok(response);

			}
			catch (ServiceException ex)
			{
				return BadRequest(ex.Message);
			}

		}

        [HttpGet]
        [Route("GetUsedDocuments")]
        public async Task<IActionResult> GetUsedDocumentsAsync()
        {
            try
            {
                var files = await _graphDriveClient.GetUsedDocumentsAsync();
                List<IUsedDocuments> response = files != null && files.Count > 0 ? DocumentMapper.RecentlyUsedDocumentMapper(files) : new List<IUsedDocuments>();                
                return Ok(response);

            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }

        }

        [HttpGet]
        [Route("GetOneDriveItems")]
        public async Task<IActionResult> GetOneDriveItemsAsync()
        {
            try
            {
                var files = await _graphDriveClient.GetOneDriveItemsAsync();
                List<IOneDriveItems> response = DocumentMapper.OneDriveItemsMapper(files.Cast<DriveItem>().ToList());
                return Ok(response);

            }
            catch (ServiceException ex)
            {
                return BadRequest(ex.Message);
            }

        }

		[HttpGet]
		[Route("GetFileStats")]
		public async Task<IActionResult> GetFileStatsAsync()
		{
			try
			{
				var imagesPattern = @"\.jpg$|\.jpeg$|\.jpe$|\.jif$|\.jfif$|\.jfi$|\.webp$|\.gif$|\.png$|\.apng$|\.bmp$|\.dib$|\.tiff$|\.tif$|\.svg$|\.svgz$|\.ico$|\.xbm$";
				var officeDocPattern = @"\.doc|\.docx$|\.dot$|\.dotx$|\.xls$|\.xlsx$|\.ppt$|\.pptx$|\.one$|\.vsd$|\.vtx$|\.vst$";
				var files = await _graphDriveClient.GetFiles();
				var folders = files.Where(a => a.Folder != null);
				var images = files.Where(a => Regex.IsMatch(a.Name, imagesPattern));
				var officeDocuments = files.Where(a => Regex.IsMatch(a.Name, officeDocPattern));
				var otherFiles = files.Except(folders).Except(images).Except(officeDocuments);

				var stats = new FileStats
				{
					TotalFoldersCount = folders.Count(),
					TotalImagesCount = images.Count(),
					TotalOfficeDocumentsCount = officeDocuments.Count(),
					TotalOthersCount = otherFiles.Count()
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
