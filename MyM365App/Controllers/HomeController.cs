using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using MyM365App.Graph;
using MyM365App.Models.Graph;
using MyM365App.Models;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Web;
using MyM365App.Services;
using System.Threading;
using MyM365App.ViewModels;
using MyM365App.Helpers;
using MyM365App.Mapper;

namespace MyM365App.Controllers
{
	[AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
	public class HomeController : Controller
	{
		private readonly ILogger<HomeController> _logger;
		private readonly GraphServices _graphServices;

		public string UserDisplayName { get; private set; } = "";
		public string UserPhoto { get; private set; } = "";
		private string UserEmail = "";
		readonly ITokenAcquisition _tokenAcquisition;

		public HomeController(ILogger<HomeController> logger, GraphServices graphServices, ITokenAcquisition tokenAcquisition)
		{
			_logger = logger;
			_graphServices = graphServices;
			_tokenAcquisition = tokenAcquisition;
		}

		public async Task<IActionResult> Index()
		{
			await GetUserInfoAsync();			
			var chat = await _graphServices.ChatClient.GetMyChats();
            ViewData["DisplayName"] = UserDisplayName;
            ViewData["ProfilePhoto"] = UserPhoto;
			ViewData["UserEmail"] = UserEmail;          	
			return View();
		}

		public IActionResult Privacy()
		{
			return View();
		}

		[AllowAnonymous]
		[ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
		public IActionResult Error()
		{
			return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
		}

		public async Task GetUserInfoAsync()
		{
			//var files = await _graphServices.DriveClient.GetCurrentUserRecentDocuments();
			//var recentFiles =  files != null  && files.Count > 10 ? files.Take(10) : null;
			var files = await _graphServices.DriveClient.GetRecentDocumentsAsync();
			var recentFiles =  files != null  && files.Count > 10 ? files.Take(10) : null;
			var user = await _graphServices.ProfileClient.GetUserProfile();
			UserDisplayName = user.DisplayName;
			UserPhoto = await _graphServices.ProfileClient.GetUserProfileImage();
			UserEmail = user.Mail;
		}
			 

        public async Task OnGetAccessTokenAsync()
		{
			// Acquire the access token.
			string[] scopes = new string[] { "user.read" };
			string accessToken = await _tokenAcquisition.GetAccessTokenForUserAsync(scopes);
			_logger.LogInformation($"Token: {accessToken}");

			// Use the access token to call a protected web API.
			HttpClient client = new HttpClient();
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
			string json = await client.GetStringAsync("https://graph.microsoft.com/v1.0/me?$select=displayName");
			_logger.LogInformation(json);
		}
	}
}