namespace MyM365App.Helpers
{
    public class Utilities
    {
        public static string GetLocalPath(string url)
        {
            string result = "";
            if (!string.IsNullOrEmpty(url))
            {
                var URL = new Uri(url);
                result = URL.AbsoluteUri;
            }
            return result;
        }
    }
}
