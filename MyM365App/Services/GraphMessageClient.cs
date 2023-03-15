using Azure.Core;
using Microsoft.Graph;
using System.Text.Json;

namespace MyM365App.Graph
{
	public class GraphMessageClient
	{
		private readonly ILogger<GraphMessageClient> _logger;
		private readonly GraphServiceClient _graphServiceClient;
		public GraphMessageClient(ILogger<GraphMessageClient> logger, GraphServiceClient graphServiceClient)
		{
			_logger = logger;
			_graphServiceClient = graphServiceClient;
		}

		public async Task<Int32> GetMessagesCountAsync(string filter)
		{
			Int32 messagesCount = 0;
			try
			{
				var options = new List<QueryOption>
				{
					 new QueryOption("$count", "true")
				};

				var results = await _graphServiceClient.Me.Messages.Request(options).Filter(filter).Select("subject").Top(1).GetAsync();

				if (results.AdditionalData.TryGetValue("@odata.count", out var propertyValue))
				{
					var propertyElement = (JsonElement)propertyValue;
					messagesCount = propertyElement.Deserialize<Int32>();
				}
			}
			catch (Exception ex)
			{

			}
			return messagesCount;
		}

		public async Task<Int32> GetMessagesCountBySearchAsync(string query)
		{
			Int32 messagesCount = 0;
			try
			{
				var options = new List<QueryOption>
				{
					new QueryOption("$search", query)

				};

				var results = await _graphServiceClient.Me.Messages.Request(options).Select("subject").Top(100).GetAsync();
                
                var pageIterator = PageIterator<Message>
                    .CreatePageIterator(
                        _graphServiceClient,
                        results,
                        // Callback executed for each item in
                        // the collection
                        (item) =>
                        {
							messagesCount += 1;
                            return true;
                        },
                        // Used to configure subsequent page
                        // requests
                        (req) =>
                        {
                            // Re-add the header to subsequent requests
                            //req.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                            return req;
                        }
                    );

                await pageIterator.IterateAsync();
            }
			catch (Exception ex)
			{

			}
			return messagesCount;
		}

	}
}

