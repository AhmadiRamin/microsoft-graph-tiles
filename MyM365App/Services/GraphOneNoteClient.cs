using Microsoft.Graph;

namespace MyM365App.Graph
{
    public class GraphOneNoteClient
    {
        private readonly ILogger<GraphOneNoteClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;
        public GraphOneNoteClient(ILogger<GraphOneNoteClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<List<Notebook>> GetNoteBooksAsync()
        {
            List<Notebook> items = new List<Notebook>();
            try
            {
                var results = await _graphServiceClient.Me.Onenote.Notebooks.Request().Top(10).GetAsync();
                items = results.ToList();
            }
            catch (Exception ex)
            {

            }
            return items;
        }

       
    }
}

