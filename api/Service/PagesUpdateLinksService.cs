using CallContent.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Net;
using System.Text.RegularExpressions;
using System.Text;

namespace CallContent.Service
{
    public class PagesUpdateLinksService
    {

        private readonly string _baseUrl = "https://graph.microsoft.com/beta/sites/";
        private readonly TokenAccessService _token;


        public PagesUpdateLinksService(TokenAccessService token)
        {
            _token = token;
        }

        public async Task<JObject> SendRequestAndParseJson(string url)
        {
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            string json = await SendRequestGraph(request);
            return JObject.Parse(json);
        }

        public async Task<string> SendRequestGraph(HttpRequestMessage request)
        {
            var token = await _token.GetTokenGraph();

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

        public async Task<List<LinkInfo>> GetContentInterns(string sharePointPageId)
        {
            string url = $"https://graph.microsoft.com/beta/sites/4156c839-562e-4702-b7ac-00c97ee6b4a8/pages/{sharePointPageId}/microsoft.graph.sitePage?expand=canvaslayout";

            List<LinkInfo> list = new List<LinkInfo>();

            JObject jObject = await SendRequestAndParseJson(url);

            string webUrl = (string)jObject["webUrl"]!;

            string pageSharePointId = GetSharepointItem(webUrl).Result!;

            string pageConfluenceId = GetConfluencePageId(pageSharePointId).Result;

            string value = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]!["value"]?[0]! ?? "";

            string innerHtml = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]!["innerHtml"]! ?? "";

            string pageTitle = (string)jObject["title"]!;

            ExtractLinksHashtag(value, innerHtml, pageTitle, pageConfluenceId, webUrl);

            return list;
        }

        public async Task<string> GetConfluencePageId(string itemId)
        {
            string url = $"{_baseUrl}4156c839-562e-4702-b7ac-00c97ee6b4a8/lists/153be7b7-ce43-4607-804f-e3773637e297/items/{itemId}?expand=fields";

            JObject jObject = await SendRequestAndParseJson(url);

            string pageId = jObject["fields"]?["pageId"]?.ToString() ?? "";

            return pageId;
        }

        // Mudar a função: pois a coluna ID já existe no Science e não precisaremos coletar o pageId interno!!
        public async Task<string?> GetWebUrlConfluencePageById(string pageId)
        {
            string baseUrl = $"{_baseUrl}4156c839-562e-4702-b7ac-00c97ee6b4a8/lists/153be7b7-ce43-4607-804f-e3773637e297/items";

            JObject jObjectItems = await SendRequestAndParseJson(baseUrl);

            foreach (var item in jObjectItems["value"] ?? new JArray())
            {
                string itemId = item["id"]?.ToString()!;

                if (string.IsNullOrEmpty(itemId))
                    continue;

                string urlItem = $"{baseUrl}/{itemId}";
                JObject itemDetails = await SendRequestAndParseJson(urlItem);

                if (itemDetails["fields"]!["pageId"]!.ToString() == pageId)
                {
                    return itemDetails["webUrl"]?.ToString();
                }
            }

            return null;
        }



        public async Task<string?> GetSharepointItem(string webUrl)
        {
            string url = _baseUrl + "4156c839-562e-4702-b7ac-00c97ee6b4a8/lists/153be7b7-ce43-4607-804f-e3773637e297/items";

            JObject jObject = await SendRequestAndParseJson(url);

            string? sharepointItemId = jObject["value"]!
                .FirstOrDefault(x => x["webUrl"]!.ToString() == webUrl)?["id"]?.ToString();

            return sharepointItemId;
        }

        public List<LinkInfo> ExtractLinksHashtag(string value, string innerHtml, string pageTitle, string pageConfluenceId, string webUrl)
        {
            var matches = Regex.Matches(innerHtml, "<a href=\"([^\"]*)\">(.*?)</a>", RegexOptions.IgnoreCase);

            List<LinkInfo> listLinks = new List<LinkInfo>();

            LogSuccessOrFail(matches, pageTitle, pageConfluenceId, webUrl);

            return listLinks;
        }

        public static void LogSuccessOrFail(MatchCollection? matches, string pageTitle, string pageConfluenceId, string webUrl)
        {
            string logFilePath = "LinkLog.txt";

            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                foreach (Match match in matches!)
                {
                    var link = match.Groups[1].Value;

                    if (!link.Contains("#"))
                    {
                        var name = Regex.Replace(match.Groups[2].Value, "<.*?>", "");
                        writer.WriteLine($"######### {pageTitle} #########");
                        writer.WriteLine($"Nome: {name}");
                        writer.WriteLine($"Link: {link}");
                        writer.WriteLine($"Status: FALHA");
                        writer.WriteLine($"Página Confluence: {pageConfluenceId}");
                        writer.WriteLine("############################");
                        writer.WriteLine("");
                        writer.WriteLine("");
                        continue;
                    }

                    var indexHashtag = link.IndexOf("#");
                    var processedLink = $"{webUrl}{link.Substring(indexHashtag)}";
                    var validName = Regex.Replace(match.Groups[2].Value, "<.*?>", "");



                    writer.WriteLine($"######### {pageTitle} #########");
                    writer.WriteLine($"Nome: {validName}");
                    writer.WriteLine($"Link: {processedLink}");
                    writer.WriteLine($"Status: SUCESSO");
                    writer.WriteLine($"Página Confluence: {pageConfluenceId}");
                    writer.WriteLine("############################");
                    writer.WriteLine("");
                    writer.WriteLine("");
                }
            }
        }

        public void ExtractExternalLinksWithPageId(string innerHtml, string pageTitle, string pageConfluenceId)
        {
            var matches = Regex.Matches(innerHtml, "<a href=\"([^\"]*)\">(.*?)</a>", RegexOptions.IgnoreCase);
            string logFilePath = "ExternalPageIdLinkLog.txt";

            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                foreach (Match match in matches)
                {
                    var link = match.Groups[1].Value;

                    if (link.Contains("pageId="))
                    {
                        var name = Regex.Replace(match.Groups[2].Value, "<.*?>", "");

                        writer.WriteLine($"######### {pageTitle} #########");
                        writer.WriteLine($"Nome: {name}");
                        writer.WriteLine($"Link: {link}");
                        writer.WriteLine($"Status: EXTERNO (com pageId)");
                        writer.WriteLine($"Página Confluence: {pageConfluenceId}");
                        writer.WriteLine("############################");
                        writer.WriteLine("");
                        writer.WriteLine("");
                    }
                }
            }
        }

        public async Task<HttpResponseMessage> UpdateSessionLinks(string sharepointPageId)
        {
            string url = $"https://graph.microsoft.com/beta/sites/4156c839-562e-4702-b7ac-00c97ee6b4a8/pages/{sharepointPageId}/microsoft.graph.sitePage?expand=canvaslayout";
            JObject jObject = await SendRequestAndParseJson(url);

            string innerHtml = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]?["innerHtml"]! ?? "";

            string webPartId = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]?["id"]! ?? "";

            string webUrl = (string)jObject["webUrl"]!;

            string pattern = @"href=""([^""]*)""";
            Regex regex = new Regex(pattern);

            var matches = regex.Matches(innerHtml);

            foreach (Match match in matches)
            {
                string link = match.Groups[1].Value;

                if (!link.Contains("#") && link.Contains("pageId=")) continue;

                int indexHashtag = link.IndexOf("#");
                string processedLink = $"{webUrl}{link.Substring(indexHashtag)}";

                innerHtml = innerHtml.Replace(link, processedLink);
            }

            var patchData = new WebPartPatch
            {
                ODataType = "#microsoft.graph.textWebPart",
                id = webPartId,
                innerHtml = innerHtml
            };

            string jsonContent = JsonConvert.SerializeObject(patchData);

            var request = new HttpRequestMessage(HttpMethod.Patch, $"https://graph.microsoft.com/beta/sites/4156c839-562e-4702-b7ac-00c97ee6b4a8/pages/{sharepointPageId}/microsoft.graph.sitePage/webParts/{webPartId}")
            {
                Content = new StringContent(jsonContent, Encoding.UTF8, "application/json")
            };

            var token = await _token.GetTokenGraph();

            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);
            HttpResponseMessage response = await client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine("Links atualizados com sucesso!");
                return response;
            }
            else
            {
                Console.WriteLine($"Erro ao atualizar links: {response.ReasonPhrase}");
                return response;
            }
        }

        public async Task<HttpResponseMessage> UpdateExternLinks(string sharepointPageId)
        {
            string url = $"https://graph.microsoft.com/beta/sites/4156c839-562e-4702-b7ac-00c97ee6b4a8/pages/{sharepointPageId}/microsoft.graph.sitePage?expand=canvaslayout";
            JObject jObject = await SendRequestAndParseJson(url);

            string innerHtml = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]?["innerHtml"]! ?? "";
            string webPartId = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]?["id"]! ?? "";

            string pattern = @"href=""([^""]*)""";
            Regex regex = new Regex(pattern);

            var matches = regex.Matches(innerHtml);

            foreach (Match match in matches)
            {
                string link = match.Groups[1].Value;

                string pageIdPattern = @"[?&]pageId=([^&]+)";
                Match pageIdMatch = Regex.Match(link, pageIdPattern);

                if (link.Contains("pageId=") && pageIdMatch.Success)
                {
                    string pageId = pageIdMatch.Groups[1].Value;
                    var webUrlByPageId = await GetWebUrlConfluencePageById(pageId);

                    string processedLink = link.Contains("#")
                        ? $"{webUrlByPageId}{link.Substring(link.IndexOf("#"))}"
                        : webUrlByPageId!;

                    innerHtml = innerHtml.Replace(link, processedLink);
                }
            }

            var patchData = new WebPartPatch
            {
                ODataType = "#microsoft.graph.textWebPart",
                id = webPartId,
                innerHtml = innerHtml
            };

            string jsonContent = JsonConvert.SerializeObject(patchData);
            var request = new HttpRequestMessage(HttpMethod.Patch, $"https://graph.microsoft.com/beta/sites/4156c839-562e-4702-b7ac-00c97ee6b4a8/pages/{sharepointPageId}/microsoft.graph.sitePage/webParts/{webPartId}")
            {
                Content = new StringContent(jsonContent, Encoding.UTF8, "application/json")
            };

            var token = await _token.GetTokenGraph();

            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);
            HttpResponseMessage response = await client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine("Links atualizados com sucesso!");
                return response;
            }
            else
            {
                Console.WriteLine($"Erro ao atualizar links: {response.ReasonPhrase}");
                return response;
            }
        }

    }
}
