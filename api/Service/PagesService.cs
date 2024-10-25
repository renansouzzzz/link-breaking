using CallContent.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Net;

namespace CallContent.Service
{
    public class PagesService
    {
        private readonly string _baseUrl = "https://graph.microsoft.com/beta/sites/";
        private readonly TokenAccessService _token;

        public PagesService(TokenAccessService token)
        {
            _token = token;
        }

        public async Task<List<SitePage>> ListPages(string siteId)
        {
            string url = _baseUrl + siteId + "/pages";

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);

            string json = await SendRequest(request);

            JObject jObject = JObject.Parse(json);

            List<SitePage> pages = jObject["value"]!.ToObject<List<SitePage>>()!;

            return pages!;
        }

        private async Task<string> SendRequest(HttpRequestMessage request)
        {
            var token = await _token.GetToken();
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);

            HttpClient http = new HttpClient();

            http.DefaultRequestHeaders.Add("Accept", "application/json;odata.metadata=none");

            var response = await http.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                string error = await response.Content.ReadAsStringAsync();
                object formatted = JsonConvert.DeserializeObject(error)!;
                throw new WebException("Error Calling the Graph API: \n" + JsonConvert.SerializeObject(formatted, Formatting.Indented));
            }

            string json = await response.Content.ReadAsStringAsync();
            return json;
        }
    }
}
