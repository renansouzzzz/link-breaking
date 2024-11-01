using CallContent.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Net;
using System.Text.RegularExpressions;
using OfficeOpenXml;

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

        public async Task<JObject> SendRequestAndParseJson(string url)
        {
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            string json = await SendRequestGraph(request);
            return JObject.Parse(json);
        }


        public async Task<List<LinkInfo>> GetContent(string sharepointPageId)
        {
            string url = $"https://graph.microsoft.com/beta/sites/4156c839-562e-4702-b7ac-00c97ee6b4a8/pages/{sharepointPageId}/microsoft.graph.sitePage?expand=canvaslayout";

            JObject jObject = await SendRequestAndParseJson(url);

            string webUrl = (string)jObject["webUrl"]!;

            string pageSharePointId = GetSharepointItem(webUrl).Result!;

            string pageConfluenceId = GetConfluencePageId(pageSharePointId).Result;

            string innerHtml = (string)jObject["canvasLayout"]?["horizontalSections"]?[0]?["columns"]?[0]?["webparts"]?[0]?["innerHtml"]! ?? "";

            string pageTitle = (string)jObject["title"]!;

            var links = ExtractLinks(innerHtml, pageTitle, pageConfluenceId);

            return links;
        }

        public async Task<List<string>> GetSharepointPageId()
        {
            string url = _baseUrl + "4156c839-562e-4702-b7ac-00c97ee6b4a8/pages";

            JObject jObject = await SendRequestAndParseJson(url);

            List<string> sharepointPageIds = jObject["value"]!
                .Select(page => page["id"]?.ToString())
                .Where(id => !string.IsNullOrEmpty(id))
                .ToList()!;

            return sharepointPageIds;
        }

        public async Task<string?> GetSharepointItem(string webUrl)
        {
            string url = _baseUrl + "4156c839-562e-4702-b7ac-00c97ee6b4a8/lists/153be7b7-ce43-4607-804f-e3773637e297/items";

            JObject jObject = await SendRequestAndParseJson(url);

            string? sharepointItemId = jObject["value"]!
                .FirstOrDefault(x => x["webUrl"]!.ToString() == webUrl)?["id"]?.ToString();

            return sharepointItemId;
        }


        public async Task<string> GetConfluencePageId(string itemId)
        {
            string url = $"{_baseUrl}4156c839-562e-4702-b7ac-00c97ee6b4a8/lists/153be7b7-ce43-4607-804f-e3773637e297/items/{itemId}?expand=fields";

            JObject jObject = await SendRequestAndParseJson(url);

            string pageId = jObject["fields"]?["pageId"]?.ToString() ?? "";

            return pageId;
        }

        private async Task<string> SendRequestGraph(HttpRequestMessage request)
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

        private List<LinkInfo> ExtractLinks(string innerHtml, string pageTitle, string pageConfluenceId)
        {
            var matches = Regex.Matches(innerHtml, "<a href=\"([^\"]*)\">(.*?)</a>", RegexOptions.IgnoreCase);

            var links = new List<LinkInfo>();

            foreach (Match match in matches)
            {
                var link = match.Groups[1].Value;
                var name = Regex.Replace(match.Groups[2].Value, "<.*?>", "");


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


        public async Task SaveLinksToExcelAsync(List<LinkInfo> linkInfos, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Links");

                int row = 1;

                var groupedLinks = linkInfos.GroupBy(link => link.PageTitle);

                foreach (var group in groupedLinks)
                {
                    foreach (var link in group)
                    {
                        worksheet.Cells[row, 1].Value = link.PageTitle;
                        worksheet.Cells[row, 2].Value = link.LinkTitle;
                        worksheet.Cells[row, 3].Value = link.LinkUrl;
                        worksheet.Cells[row, 4].Value = link.PageId;
                        row++;
                    }
                }

                FileInfo excelFile = new FileInfo(filePath);
                await package.SaveAsAsync(excelFile);
            }
        }
    }
}
