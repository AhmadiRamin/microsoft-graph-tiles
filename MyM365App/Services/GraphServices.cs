using Microsoft.Graph;
using MyM365App.Graph;

namespace MyM365App.Services
{
	public class GraphServices
	{
        public GraphDriveClient DriveClient { get; set; }
        public GraphProfileClient ProfileClient { get; set; }
        public GraphMessageClient MessageClient { get; set; }
        public GraphOneNoteClient OneNoteClient { get; set; }
        public GraphChatClient ChatClient { get; set; }

        public GraphServices(
			GraphMessageClient messageClient,
			GraphDriveClient driveClient,
			GraphProfileClient profileClient,
			GraphOneNoteClient oneNoteClient,
			GraphChatClient chatClient

            )
		{
			DriveClient = driveClient;
			MessageClient = messageClient;
			ProfileClient = profileClient;
			OneNoteClient = oneNoteClient;
			ChatClient = chatClient;
        }
	}
}
