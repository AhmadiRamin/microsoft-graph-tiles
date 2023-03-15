using System.Net;
using System.Net.Http.Headers;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Primitives;
using Microsoft.Graph;
using MyM365App.Graph;

namespace MyM365App.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class ProxyController : ControllerBase
	{
		private readonly ILogger<GraphProfileClient> _logger;
		private readonly GraphServiceClient _graphServiceClient;

		public ProxyController(ILogger<GraphProfileClient> logger, GraphServiceClient graphServiceClient)
		{
			_logger = logger;
			_graphServiceClient = graphServiceClient;
		}

		[HttpGet]
		[Route("{*all}")]
		public async Task<IActionResult> GetAsync(string all)
		{
			return await ProcessRequestAsync(Microsoft.Graph.HttpMethods.GET, all, null).ConfigureAwait(false);
		}

		[HttpPost]
		[Route("{*all}")]
		public async Task<IActionResult> PostAsync(string all, [FromBody] object body)
		{
			return await ProcessRequestAsync(Microsoft.Graph.HttpMethods.POST, all, body).ConfigureAwait(false);
		}

		[HttpDelete]
		[Route("{*all}")]
		public async Task<IActionResult> DeleteAsync(string all)
		{
			return await ProcessRequestAsync(Microsoft.Graph.HttpMethods.DELETE, all, null).ConfigureAwait(false);
		}

		[HttpPut]
		[Route("{*all}")]
		public async Task<IActionResult> PutAsync(string all, [FromBody] object body)
		{
			return await ProcessRequestAsync(Microsoft.Graph.HttpMethods.PUT, all, body).ConfigureAwait(false);
		}

		[HttpPatch]
		[Route("{*all}")]
		public async Task<IActionResult> PatchAsync(string all, [FromBody] object body)
		{
			return await ProcessRequestAsync(Microsoft.Graph.HttpMethods.PATCH, all, body).ConfigureAwait(false);
		}

		private async Task<IActionResult> ProcessRequestAsync(Microsoft.Graph.HttpMethods method, string all, object content)
		{
			
			var qs = HttpContext.Request.QueryString;
			var url = $"{GetBaseUrlWithoutVersion(_graphServiceClient)}/{all}{qs.ToUriComponent()}";

			var request = new BaseRequest(url, _graphServiceClient, null)
			{
				Method = method,
				ContentType = HttpContext.Request.ContentType,
			};

			var neededHeaders = Request.Headers.Where(h => h.Key.ToLower() == "if-match" || h.Key.ToLower() == "consistencylevel").ToList();
			if (neededHeaders.Count() > 0)
			{
				foreach (var header in neededHeaders)
				{
					request.Headers.Add(new HeaderOption(header.Key, string.Join(",", header.Value)));
				}
			}

			var contentType = "application/json";

			try
			{
				using (var response = await request.SendRequestAsync(content?.ToString(), CancellationToken.None, HttpCompletionOption.ResponseContentRead).ConfigureAwait(false))
				{
					response.Content.Headers.TryGetValues("content-type", out var contentTypes);

					contentType = contentTypes?.FirstOrDefault() ?? contentType;

					var byteArrayContent = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
					return new HttpResponseMessageResult(ReturnHttpResponseMessage(HttpStatusCode.OK, contentType, new ByteArrayContent(byteArrayContent)));
				}
			}
			catch (ServiceException ex)
			{
				return new HttpResponseMessageResult(ReturnHttpResponseMessage(ex.StatusCode, contentType, new StringContent(ex.Error.ToString())));
			}
		}

		private static HttpResponseMessage ReturnHttpResponseMessage(HttpStatusCode httpStatusCode, string contentType, HttpContent httpContent)
		{
			var httpResponseMessage = new HttpResponseMessage(httpStatusCode)
			{
				Content = httpContent
			};

			try
			{
				httpResponseMessage.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
			}
			catch
			{
				httpResponseMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
			}

			return httpResponseMessage;
		}

		private string GetBaseUrlWithoutVersion(GraphServiceClient graphClient)
		{
			var baseUrl = graphClient.BaseUrl;
			var index = baseUrl.LastIndexOf('/');
			return baseUrl.Substring(0, index);
		}

		public class HttpResponseMessageResult : IActionResult
		{
			private readonly HttpResponseMessage _responseMessage;

			public HttpResponseMessageResult(HttpResponseMessage responseMessage)
			{
				_responseMessage = responseMessage; // could add throw if null
			}

			public async Task ExecuteResultAsync(ActionContext context)
			{
				context.HttpContext.Response.StatusCode = (int)_responseMessage.StatusCode;

				foreach (var header in _responseMessage.Headers)
				{
					context.HttpContext.Response.Headers.TryAdd(header.Key, new StringValues(header.Value.ToArray()));
				}

				context.HttpContext.Response.ContentType = _responseMessage.Content.Headers.ContentType.ToString();

				using (var stream = await _responseMessage.Content.ReadAsStreamAsync())
				{
					await stream.CopyToAsync(context.HttpContext.Response.Body);
					await context.HttpContext.Response.Body.FlushAsync();
				}
			}
		}

	}
}
