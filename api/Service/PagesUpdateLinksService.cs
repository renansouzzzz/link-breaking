using CallContent.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Net;
using System.Text.RegularExpressions;
using System.Collections.Generic;

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

            string innerHtml = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]?["innerHtml"]! ?? "";

            string pageTitle = (string)jObject["title"]!;

            ExtractLinksHashtag(innerHtml, pageTitle, pageConfluenceId, webUrl);

            return list;
        }

        public async Task<List<LinkInfo>> GetContentExtern(string sharepointPageId)
        {
            string url = $"https://graph.microsoft.com/beta/sites/4156c839-562e-4702-b7ac-00c97ee6b4a8/pages/{sharepointPageId}/microsoft.graph.sitePage?expand=canvaslayout";

            JObject jObject = await SendRequestAndParseJson(url);

            string webUrl = (string)jObject["webUrl"]!;

            string pageSharePointId = GetSharepointItem(webUrl).Result!;

            string pageConfluenceId = GetConfluencePageId(pageSharePointId).Result;

            string innerHtml = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]?["innerHtml"]! ?? "";

            string pageTitle = (string)jObject["title"]!;

            var extractLinks = ExtractLinksNoHashtag(innerHtml, pageTitle, pageConfluenceId);

            return extractLinks;
        }

        public async Task<string> GetConfluencePageId(string itemId)
        {
            string url = $"{_baseUrl}4156c839-562e-4702-b7ac-00c97ee6b4a8/lists/153be7b7-ce43-4607-804f-e3773637e297/items/{itemId}?expand=fields";

            JObject jObject = await SendRequestAndParseJson(url);

            string pageId = jObject["fields"]?["pageId"]?.ToString() ?? "";

            return pageId;
        }

        public async Task<string?> GetSharepointItem(string webUrl)
        {
            string url = _baseUrl + "4156c839-562e-4702-b7ac-00c97ee6b4a8/lists/153be7b7-ce43-4607-804f-e3773637e297/items";

            JObject jObject = await SendRequestAndParseJson(url);

            string? sharepointItemId = jObject["value"]!
                .FirstOrDefault(x => x["webUrl"]!.ToString() == webUrl)?["id"]?.ToString();

            return sharepointItemId;
        }

        public List<LinkInfo> ExtractLinksHashtag(string innerHtml, string pageTitle, string pageConfluenceId, string webUrl)
        {
            var matches = Regex.Matches(innerHtml, "<a href=\"([^\"]*)\">(.*?)</a>", RegexOptions.IgnoreCase);

            List<LinkInfo> listLinks = new List<LinkInfo>();

            string logFilePath = "LinkLog.txt";

            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                foreach (Match match in matches)
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

                    //listLinks.Add(processedLink);

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

            return listLinks;
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

        public List<LinkInfo> ExtractLinksNoHashtag(string innerHtml, string pageTitle, string pageConfluenceId)
        {
            var matches = Regex.Matches(innerHtml, "<a href=\"([^\"]*)\">(.*?)</a>", RegexOptions.IgnoreCase);

            var links = new List<LinkInfo>();

            foreach (Match match in matches)
            {
                var link = match.Groups[1].Value;
                var name = Regex.Replace(match.Groups[2].Value, "<.*?>", "");

                if (!link.Contains("#"))

                    links.Add(new LinkInfo
                    {
                        PageTitle = pageTitle,
                        LinkTitle = name,
                        LinkUrl = link,
                        PageId = pageConfluenceId
                    });
            }

            return links;
        }
    }
}
